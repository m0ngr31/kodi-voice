#!/usr/bin/env python

# For a complete discussion, see http://forum.kodi.tv/showthread.php?tid=254502

import datetime
import threading
import json
import hashlib
import time
import codecs
import urllib
import os
import io
import random
import re
import string
import sys
import unicodedata
import logging
import roman
import requests
from num2words import num2words
from fuzzywuzzy import fuzz, process
from ConfigParser import SafeConfigParser
from .cache import KodiCache


log = logging.getLogger(__name__)

SORT_RATING = {"method": "rating", "order": "descending"}
SORT_RANDOM = {"method": "random", "order": "descending"}
SORT_YEAR = {"method": "year", "order": "descending"}
SORT_TITLE = {"method": "title", "order": "ascending"}
SORT_DATEADDED = {"method": "dateadded", "order": "descending"}
SORT_LASTPLAYED = {"method": "lastplayed", "order": "descending"}
SORT_EPISODE = {"method": "episode", "order": "ascending"}

FILTER_UNWATCHED = {"operator": "lessthan", "field": "playcount", "value": "1"}
FILTER_WATCHED = {"operator": "isnot", "field": "playcount", "value": "0"}

# For recommendations, we sort (not filter) the results on rating and then
# choose a random item from the results.  To keep the quality high, we want to
# limit the number of results before picking a random item.
LIMIT_RECOMMENDED_MOVIES = (0, 40)
LIMIT_RECOMMENDED_SHOWS = (0, 20)
LIMIT_RECOMMENDED_EPISODES = (0, 100)
LIMIT_RECOMMENDED_MUSICVIDEOS = (0, 40)
LIMIT_RECOMMENDED_ARTISTS = (0, 20)
LIMIT_RECOMMENDED_ALBUMS = (0, 40)
LIMIT_RECOMMENDED_SONGS = (0, 100)


def sanitize_name(media_name, normalize=True):
  if normalize:
    try:
      # Normalize string
      name = unicodedata.normalize('NFKD', media_name).encode('ASCII', 'ignore')
    except:
      name = media_name
  else:
    name = media_name

  # Remove invalid characters, per Amazon:
  # Slot type values can contain alphanumeric characters, spaces, commas,
  # apostrophes, periods, hyphens, ampersands and the @ symbol only.
  name = re.sub(r'[`~!#$%^*()_=+\[\]{}\\|;:"<>/?]', '', name)

  # Slot items cannot exceed 140 chars, per Amazon
  if len(name) > 140:
    name = name[:140].rsplit(' ', 1)[0]

  name = name.strip()
  return name


# Remove extra slashes
def http_normalize_slashes(url):
  url = str(url)
  segments = url.split('/')
  correct_segments = []
  for segment in segments:
    if segment != '':
      correct_segments.append(segment)
  first_segment = str(correct_segments[0])
  if first_segment.find('http') == -1:
    correct_segments = ['http:'] + correct_segments
  correct_segments[0] = correct_segments[0] + '/'
  normalized_url = '/'.join(correct_segments)
  return normalized_url


def RPCString(method, params=None, sort=None, filters=None, fields=None, limits=None, filtertype=None):
  j = {"jsonrpc": "2.0", "method": method, "id": 1}
  j["params"] = {}
  if params:
    j["params"] = params
  if sort:
    j["params"]["sort"] = sort
  if filters:
    if not filtertype:
      filtertype = "and"
    if len(filters) > 1:
      j["params"]["filter"] = {filtertype: filters}
    else:
      j["params"]["filter"] = filters[0]
  if fields:
    j["params"]["properties"] = fields
  if limits:
    j["params"]["limits"] = {"start": limits[0], "end": limits[1]}
  return json.dumps(j)


# Replace digits with word-form numbers.
def digits2words(phrase, lang='en'):
  wordified = ''
  for word in phrase.split():
    if word.isnumeric():
      word = num2words(float(word), lang=lang)
    wordified = wordified + word + " "
  return wordified[:-1]


# Replace word-form numbers with digits.
def words2digits(phrase, lang='en'):
  numwords = {}

  numwords_file = os.path.join(os.path.dirname(__file__), "NUMWORDS." + lang + ".txt")
  f = codecs.open(numwords_file, 'rb', 'utf-8')
  for line in f:
    l = line.encode("utf-8").strip().split('|')
    if l[0] == 'connectors':
      connectors = l[1:]
      for words in connectors:
        for word in words.strip().split():
          numwords[word.decode('utf-8')] = (1, 0, 0)
    if l[0] == 'units':
      units = l[1:]
      for idx, words in enumerate(units):
        for word in words.strip().split():
          numwords[word.decode('utf-8')] = (1, idx, 1)
    if l[0] == 'tens':
      tens = l[1:]
      for idx, words in enumerate(tens):
        for word in words.strip().split():
          numwords[word.decode('utf-8')] = (1, idx * 10, 2)
    if l[0] == 'scales':
      scales = l[1:]
      for idx, words in enumerate(scales):
        for word in words.strip().split():
          numwords[word.decode('utf-8')] = (10 ** (idx * 3 or 2), 0, 3)
  f.close()

  wordified = ''
  current = result = 0
  prev_level = sys.maxint
  in_number = False
  phrase = re.sub(r'[-]', ' ', phrase)
  for word in phrase.split():
    if word not in numwords:
      if in_number:
        wordified = wordified + str(current + result) + " "
        current = result = 0
        prev_level = sys.maxint
      in_number = False
      wordified = wordified + word + " "
    else:
      in_number = True
      scale, increment, level = numwords[word]

      # Handle things like "nine o two one o" (9 0 2 1 0)
      if level == prev_level == 1:
        wordified = wordified + str(current) + " "
        current = result = 0

      prev_level = level

      # account for things like "hundred fifty" vs "one hundred fifty"
      if scale >= 100 and current == 0:
        current = 1

      current = current * scale + increment
      if scale > 100:
        result += current
        current = 0

  if in_number:
    wordified = wordified + str(current + result) + " "

  return wordified[:-1]


# Replace digits with roman numerals.
def digits2roman(phrase, lang='en'):
  wordified = ''
  for word in phrase.split():
    if word.isnumeric():
      word = roman.toRoman(int(word))
    wordified = wordified + word + " "
  return wordified[:-1]


# Replace word-form numbers with roman numerals.
def words2roman(phrase, lang='en'):
  return digits2roman(words2digits(phrase, lang=lang), lang=lang)


# Provide a map from ISO code (both bibliographic and terminologic)
# in ISO 639-2 to a dict with the two letter ISO 639-2 codes (alpha2)
# English and french names
#
# "bibliographic" iso codes are derived from English word for the language
# "terminologic" iso codes are derived from the pronunciation in the target
# language (if different to the bibliographic code)
#
# Source
# http://stackoverflow.com/questions/2879856/get-system-language-in-iso-639-3-letter-codes-in-python/2879958#2879958
#
# Usage
# country_dic = getisocodes_dict()
# print country_dic['eng']
def getisocodes_dict():
  D = {}
  country_dic_file = os.path.join(os.path.dirname(__file__), "ISO-639-2_utf-8.txt")
  f = codecs.open(country_dic_file, 'rb', 'utf-8')
  for line in f:
    iD = {}
    iD['bibliographic'], iD['terminologic'], iD['alpha2'], iD['en'], iD['fr'], iD['de'] = line.encode("utf-8").strip().split('|')
    D[iD['bibliographic']] = iD

    if iD['terminologic']:
      D[iD['terminologic']] = iD

    if iD['alpha2']:
      D[iD['alpha2']] = iD

    for k in iD:
      # Assign `None` when columns not available from the data
      iD[k] = iD[k] or None
  f.close()
  return D


class KodiConfigParser(SafeConfigParser):
  def __init__(self, config_file=None, *args, **kwargs):
    SafeConfigParser.__init__(self, *args, **kwargs)

    # Seed the default values from the example
    self.config_file = os.path.join(os.path.dirname(__file__), "kodi.config.example")
    self.read(self.config_file)

    if not os.path.isfile(config_file):
      # Fill out the rest of the config based on .env variables
      SCHEME = os.getenv('KODI_SCHEME')
      if SCHEME and SCHEME != 'None':
        self.set('DEFAULT', 'scheme', SCHEME)
      SUBPATH = os.getenv('KODI_SUBPATH')
      if SUBPATH and SUBPATH != 'None':
        self.set('DEFAULT', 'subpath', SUBPATH)
      KODI_ADDRESS = os.getenv('KODI_ADDRESS')
      if KODI_ADDRESS and KODI_ADDRESS != 'None':
        self.set('DEFAULT', 'address', KODI_ADDRESS)
      KODI_PORT = os.getenv('KODI_PORT')
      if KODI_PORT and KODI_PORT != 'None':
        self.set('DEFAULT', 'port', KODI_PORT)
      KODI_USERNAME = os.getenv('KODI_USERNAME')
      if KODI_USERNAME and KODI_USERNAME != 'None':
        self.set('DEFAULT', 'username', KODI_USERNAME)
      KODI_PASSWORD = os.getenv('KODI_PASSWORD')
      if KODI_PASSWORD and KODI_PASSWORD != 'None':
        self.set('DEFAULT', 'password', KODI_PASSWORD)
      CACHE_BUCKET = os.getenv('CACHE_BUCKET')
      if CACHE_BUCKET and CACHE_BUCKET != 'None':
        self.set('DEFAULT', 'cache_bucket', CACHE_BUCKET)
      S3_CACHE_AWS_ACCESS_KEY_ID = os.getenv('S3_CACHE_AWS_ACCESS_KEY_ID')
      if S3_CACHE_AWS_ACCESS_KEY_ID and S3_CACHE_AWS_ACCESS_KEY_ID != 'None':
        self.set('DEFAULT', 's3_cache_aws_access_key_id', S3_CACHE_AWS_ACCESS_KEY_ID)
      S3_CACHE_AWS_SECRET_ACCESS_KEY = os.getenv('S3_CACHE_AWS_SECRET_ACCESS_KEY')
      if S3_CACHE_AWS_SECRET_ACCESS_KEY and S3_CACHE_AWS_SECRET_ACCESS_KEY != 'None':
        self.set('DEFAULT', 's3_cache_aws_secret_access_key', S3_CACHE_AWS_SECRET_ACCESS_KEY)
      OWNCLOUD_CACHE_URL = os.getenv('OWNCLOUD_CACHE_URL')
      if OWNCLOUD_CACHE_URL and OWNCLOUD_CACHE_URL != 'None':
        self.set('DEFAULT', 'owncloud_cache_url', OWNCLOUD_CACHE_URL)
      OWNCLOUD_CACHE_USER = os.getenv('OWNCLOUD_CACHE_USER')
      if OWNCLOUD_CACHE_USER and OWNCLOUD_CACHE_USER != 'None':
        self.set('DEFAULT', 'owncloud_cache_user', OWNCLOUD_CACHE_USER)
      OWNCLOUD_CACHE_PASSWORD = os.getenv('OWNCLOUD_CACHE_PASSWORD')
      if OWNCLOUD_CACHE_PASSWORD and OWNCLOUD_CACHE_PASSWORD != 'None':
        self.set('DEFAULT', 'owncloud_cache_password', OWNCLOUD_CACHE_PASSWORD)
      READ_TIMEOUT = os.getenv('READ_TIMEOUT')
      if READ_TIMEOUT and READ_TIMEOUT != 'None':
        self.set('DEFAULT', 'read_timeout', READ_TIMEOUT)
      READ_TIMEOUT_ASYNC = os.getenv('READ_TIMEOUT_ASYNC')
      if READ_TIMEOUT_ASYNC and READ_TIMEOUT_ASYNC != 'None':
        self.set('DEFAULT', 'read_timeout_async', READ_TIMEOUT_ASYNC)
      SHUTDOWN_MEANS_QUIT = os.getenv('SHUTDOWN_MEANS_QUIT')
      if SHUTDOWN_MEANS_QUIT and SHUTDOWN_MEANS_QUIT != 'None':
        self.set('DEFAULT', 'shutdown', SHUTDOWN_MEANS_QUIT)
      SKILL_TZ = os.getenv('SKILL_TZ')
      if SKILL_TZ and SKILL_TZ != 'None':
        self.set('DEFAULT', 'timezone', SKILL_TZ)
      ACCEPT_MUSIC_WARNING = os.getenv('ACCEPT_MUSIC_WARNING')
      if ACCEPT_MUSIC_WARNING and ACCEPT_MUSIC_WARNING != 'None':
        self.set('DEFAULT', 'accept_music_warning', ACCEPT_MUSIC_WARNING)
      USE_PROXY = os.getenv('USE_PROXY')
      if USE_PROXY and USE_PROXY != 'None':
        self.set('DEFAULT', 'use_proxy', USE_PROXY)
      ALT_PROXY = os.getenv('ALT_PROXY')
      if ALT_PROXY and ALT_PROXY != 'None':
        self.set('DEFAULT', 'alt_proxy', ALT_PROXY)
      MONGODB_URI = os.getenv('MONGODB_URI')
      if MONGODB_URI and MONGODB_URI != 'None':
        self.set('DEFAULT', 'mongodb_uri', MONGODB_URI)
      LANGUAGE = os.getenv('LANGUAGE', 'en') # Default to English
      if LANGUAGE and LANGUAGE != 'None':
        self.set('global', 'language', LANGUAGE)
      DEEP_SEARCH = os.getenv('DEEP_SEARCH')
      if DEEP_SEARCH and DEEP_SEARCH != 'None':
        self.set('global', 'deep_search', DEEP_SEARCH)
      MAX_PLAYLIST_ITEMS = os.getenv('PLAYLIST_ITEMS')
      if MAX_PLAYLIST_ITEMS and MAX_PLAYLIST_ITEMS != 'None':
        self.set('global', 'playlist_max_items', MAX_PLAYLIST_ITEMS)
      MAX_UNWATCHED_SHOWS = os.getenv('MAX_UNWATCHED_SHOWS')
      if MAX_UNWATCHED_SHOWS and MAX_UNWATCHED_SHOWS != 'None':
        self.set('global', 'unwatched_shows_max_results', MAX_UNWATCHED_SHOWS)
      MAX_UNWATCHED_EPISODES = os.getenv('MAX_UNWATCHED_EPISODES')
      if MAX_UNWATCHED_EPISODES and MAX_UNWATCHED_EPISODES != 'None':
        self.set('global', 'unwatched_episodes_max_results', MAX_UNWATCHED_EPISODES)
      MAX_UNWATCHED_MOVIES = os.getenv('MAX_UNWATCHED_MOVIES')
      if MAX_UNWATCHED_MOVIES and MAX_UNWATCHED_MOVIES != 'None':
        self.set('global', 'unwatched_movies_max_results', MAX_UNWATCHED_MOVIES)
      SKILL_LOGLEVEL = os.getenv('SKILL_LOGLEVEL')
      if SKILL_LOGLEVEL and SKILL_LOGLEVEL != 'None':
        self.set('global', 'loglevel', SKILL_LOGLEVEL)
      # unconditionally disable this as Heroku deployments don't need to see the device IDs,
      # since we don't support device mapping with Heroku.
      self.set('alexa', 'logsensitive', 'no')
      SKILL_APPID = os.getenv('SKILL_APPID')
      if SKILL_APPID and SKILL_APPID != 'None':
        self.set('alexa', 'skill_id', SKILL_APPID)
      SLOT_ITEMS_MAX = os.getenv('slot_items_max')
      if SLOT_ITEMS_MAX and SLOT_ITEMS_MAX != 'None':
        self.set('alexa', 'slot_items_max', SLOT_ITEMS_MAX)
    else:
      self.config_file = config_file
      self.read(self.config_file)


class Kodi:
  def __init__(self, config=None, context=None):
    self.config = config
    self.config_error = False

    # When testing from the web simulator there is no context object (04/2017)
    try:
      self.deviceId = context.System.device.deviceId
    except:
      self.deviceId = 'Unknown Device'

    if self.config.has_section(self.deviceId):
      self.dev_cfg_section = self.deviceId
    else:
      self.dev_cfg_section = 'DEFAULT'

    self.language = self.config.get('global', 'language').lower()
    self.playlist_limit = self.config.get('global', 'playlist_max_items')
    if self.playlist_limit and self.playlist_limit != 'None':
      self.playlist_limit = int(self.playlist_limit)
    else:
      self.playlist_limit = sys.maxint
    self.max_unwatched_shows = int(self.config.get('global', 'unwatched_shows_max_results'))
    self.max_unwatched_episodes = int(self.config.get('global', 'unwatched_episodes_max_results'))
    self.max_unwatched_movies = int(self.config.get('global', 'unwatched_movies_max_results'))
    self.logsensitive = self.config.getboolean('global', 'logsensitive')

    try:
      self.scheme = self.config.get(self.dev_cfg_section, 'scheme')
      self.subpath = self.config.get(self.dev_cfg_section, 'subpath')
      self.address = self.config.get(self.dev_cfg_section, 'address')
      self.port = self.config.get(self.dev_cfg_section, 'port')
      self.username = self.config.get(self.dev_cfg_section, 'username')
      self.password = self.config.get(self.dev_cfg_section, 'password')
      self.read_timeout = float(self.config.get(self.dev_cfg_section, 'read_timeout'))
      self.read_timeout_async = float(self.config.get(self.dev_cfg_section, 'read_timeout_async'))
    except:
      self.config_error = True

    if not self.scheme or not self.address or not self.port or not self.username or not self.password:
      self.config_error = True

    cache_bucket = self.config.get(self.dev_cfg_section, 'cache_bucket')
    if not cache_bucket or cache_bucket == 'None':
      cache_bucket = None
    s3_cache_key_id = self.config.get(self.dev_cfg_section, 's3_cache_aws_access_key_id')
    if not s3_cache_key_id or s3_cache_key_id == 'None':
      s3_cache_key_id = None
    s3_cache_key = self.config.get(self.dev_cfg_section, 's3_cache_aws_secret_access_key')
    if not s3_cache_key or s3_cache_key == 'None':
      s3_cache_key = None
    oc_cache_url = self.config.get(self.dev_cfg_section, 'owncloud_cache_url')
    if not oc_cache_url or oc_cache_url == 'None':
      oc_cache_url = None
    oc_cache_user = self.config.get(self.dev_cfg_section, 'owncloud_cache_user')
    if not oc_cache_user or oc_cache_user == 'None':
      oc_cache_user = None
    oc_cache_pass = self.config.get(self.dev_cfg_section, 'owncloud_cache_password')
    if not oc_cache_pass or oc_cache_pass == 'None':
      oc_cache_pass = None

    self.cache = KodiCache(cache_bucket,
            aws_access_key_id=s3_cache_key_id, aws_secret_access_key=s3_cache_key,
            oc_url=oc_cache_url, oc_user=oc_cache_user, oc_password=oc_cache_pass)

    try:
      # On a successful cache hit, this variable tells the skill to fetch a fresh
      # copy from Kodi in the background on a worker thread.
      #
      # This is an undocumented/hidden option because a) it provides minimal real
      # value, b) can be a source of confusion for the user, and c) doesn't work
      # on most cloud deployments because the main thread terminates before the
      # worker thread completes.
      self.cache_bg_update = self.config.getboolean(self.dev_cfg_section, 'cache_bg_update')
    except:
      self.cache_bg_update = False

  # Construct the JSON-RPC message and send it to the Kodi player
  def SendCommand(self, command, wait_resp=True, cache_resp=False):
    # Join the configuration variables into a url
    url = "%s://%s:%s/%s/%s" % (self.scheme, self.address, self.port, self.subpath, 'jsonrpc')

    # Remove any double slashes in the url
    url = http_normalize_slashes(url)

    log.info('Received request from device %s', self.deviceId if self.logsensitive else '[hidden]')
    log.info('Sending request to %s', url if self.logsensitive else '[hidden]')
    log.debug(command)

    timeout = (10, self.read_timeout)
    if not wait_resp:
      # set the read timeout (the second value here) to something really small
      # to 'fake' a non-blocking call.  we want the connect and transmit to
      # block, but just ignore the response from Kodi.
      timeout = (10, self.read_timeout_async)

    # Try to fetch from cache
    r = None
    cache_file = None
    if self.cache.enabled and cache_resp and wait_resp:
      h = hashlib.sha1()
      h.update(command)
      h.update(url)
      cache_file = h.hexdigest()
      del h
      r = self.cache.get(cache_file)

    auth = (self.username, self.password)

    if self.cache.enabled and r:
      # fetched the response from cache, so let's return it immediately but
      # update the cache object in the background.
      if self.cache_bg_update:
        t = threading.Thread(target=self.cache.add, args=(cache_file, url, auth, command, (60, 120)))
        t.daemon = True
        t.start()
      return r
    else:
      # no cached response found, so send the command directly to Kodi and,
      # if caching is enabled, cache the response.
      return self.cache.add(cache_file, url, auth, command, timeout, wait_resp)


  # Utilities

  def sanitize_name(self, *args, **kwargs):
    return sanitize_name(*args, **kwargs)


  # Helpers to find media

  # Match heard string to something in the results
  def matchHeard(self, heard, results, lookingFor='label', limit=10):
    located = []

    heard_lower = heard.lower()

    # Very ugly hack for German Alexa.  In English, if a user specifies
    # 'percent', she converts it to a '%' symbol.  In German, for whatever
    # reason, she leaves it unconverted as 'prozent'.  Let's convert here.
    heard_lower = re.sub(r'prozent(?=[.,\s]|$)', '%', heard_lower)

    log.info('Trying to match: %s', heard_lower.encode("utf-8"))

    heard_ascii = sanitize_name(heard_lower)
    for result in results:
      result_lower = result[lookingFor].lower()

      # Direct comparison
      if type(heard_lower) is type(result_lower):
        if result_lower == heard_lower:
          log.info('Simple match on direct comparison')
          located.append(result)
          continue

      # Strip out non-ascii symbols
      result_name = sanitize_name(result_lower)

      # Direct comparison (ASCII)
      if result_name == heard_ascii:
        log.info('Simple match on direct comparison (ASCII)')
        located.append(result)
        continue

    if not located:
      log.info('Simple match failed, trying fuzzy match')
      log.info('Processing %d items with fuzzywuzzy...', len(results))

      match_strings = []
      for f in (None, digits2roman, words2roman, words2digits, digits2words):
        try:
          if f is not None:
            match_string = f(heard_lower, self.language)
            match_func = f.__name__
          else:
            match_string = heard_lower
            match_func = 'heard'

          match_strings.append(match_string)
          log.info('  %s -> "%s"', match_func, match_string.encode("utf-8"))
        except:
          continue

      fuzzy_results = []
      for ms in set(match_strings):
        log.info('  Trying with "%s"', ms.encode("utf-8"))
        matches = process.extractBests(ms, [d[lookingFor] for d in results], limit=limit, scorer=fuzz.UQRatio, score_cutoff=75)
        if matches:
          log.info('    Best score %d%%', matches[0][1])
          fuzzy_results += matches

      # Got a match?
      if fuzzy_results:
        winners = sorted(fuzzy_results, key=lambda x: x[1], reverse=True)
        log.info('BEST MATCH: "%s" @ %d%%', winners[0][0].encode("utf-8"), winners[0][1])
        for winner in winners:
          located.append((item for item in results if item[lookingFor] == winner[0]).next())
    else:
      log.info('BEST MATCH: "%s"', located[0][lookingFor].encode("utf-8"))

    return located[:limit]


  def FindVideoPlaylist(self, heard_search):
    log.info('Searching for video playlist "%s"', heard_search.encode("utf-8"))

    located = []
    playlists = self.GetVideoPlaylists()
    if 'result' in playlists and 'files' in playlists['result']:
      ll = self.matchHeard(heard_search, playlists['result']['files'])
      if ll:
        located = [(item['file'], item['label']) for item in ll]

    return located

  def FindAudioPlaylist(self, heard_search):
    log.info('Searching for audio playlist "%s"', heard_search.encode("utf-8"))

    located = []
    playlists = self.GetMusicPlaylists()
    if 'result' in playlists and 'files' in playlists['result']:
      ll = self.matchHeard(heard_search, playlists['result']['files'])
      if ll:
        located = [(item['file'], item['label']) for item in ll]

    return located

  def FindVideoGenre(self, heard_search, genretype='movie'):
    log.info('Searching for %s genre "%s"', genretype, heard_search.encode("utf-8"))

    located = []
    genres = self.GetVideoGenres(genretype)
    if 'result' in genres and 'genres' in genres['result']:
      ll = self.matchHeard(heard_search, genres['result']['genres'])
      if ll:
        located = [(item['genreid'], item['label']) for item in ll]

    return located

  def FindMovie(self, heard_search):
    log.info('Searching for movie "%s"', heard_search.encode("utf-8"))

    located = []
    movies = self.GetMovies()
    if 'result' in movies and 'movies' in movies['result']:
      ll = self.matchHeard(heard_search, movies['result']['movies'])
      if ll:
        located = [(item['movieid'], item['label']) for item in ll]

    return located

  def FindTvShow(self, heard_search):
    log.info('Searching for show "%s"', heard_search.encode("utf-8"))

    located = []
    shows = self.GetShows()
    if 'result' in shows and 'tvshows' in shows['result']:
      ll = self.matchHeard(heard_search, shows['result']['tvshows'])
      if ll:
        located = [(item['tvshowid'], item['label']) for item in ll]

    return located

  # There is no JSON-RPC method for VideoLibrary.GetArtists, so we need a way
  # to filter the library results here.
  def FilterMusicVideosByArtist(self, results, artist):
    log.info('Searching for music videos by "%s"', artist.encode("utf-8"))

    # Kodi.matchHeard() expects to match on strings, but Kodi gives us arrays
    # for the artist fields.  I'm not entirely sure, but I presume it's for
    # alternate artist names.  For simplicity (and until someone complains),
    # let's just choose the first artist label to match on.
    artistvideos = [{k: (v if k != u'artist' else v[0]) for k, v in d.items()} for d in results]
    return self.matchHeard(artist, artistvideos, 'artist', sys.maxint)

  def FindMusicVideo(self, heard_search, heard_artist=None):
    log.info('Searching for music video "%s"', heard_search.encode("utf-8"))

    located = []
    mvs = self.GetMusicVideos()
    if 'result' in mvs and 'musicvideos' in mvs['result']:
      if heard_artist:
        musicvideos = self.FilterMusicVideosByArtist(mvs['result']['musicvideos'], heard_artist)
      else:
        musicvideos = mvs['result']['musicvideos']
      ll = self.matchHeard(heard_search, musicvideos)
      if ll:
        located = [(item['musicvideoid'], item['label']) for item in ll]

    return located

  def FindMusicGenre(self, heard_search):
    log.info('Searching for music genre "%s"', heard_search.encode("utf-8"))

    located = []
    genres = self.GetMusicGenres()
    if 'result' in genres and 'genres' in genres['result']:
      ll = self.matchHeard(heard_search, genres['result']['genres'])
      if ll:
        located = [(item['genreid'], item['label']) for item in ll]

    return located

  def FindArtist(self, heard_search):
    log.info('Searching for artist "%s"', heard_search.encode("utf-8"))

    located = []
    artists = self.GetMusicArtists()
    if 'result' in artists and 'artists' in artists['result']:
      ll = self.matchHeard(heard_search, artists['result']['artists'], 'artist')
      if ll:
        located = [(item['artistid'], item['label']) for item in ll]

    return located

  def FindAlbum(self, heard_search, artist_id=None):
    log.info('Searching for album "%s"', heard_search.encode("utf-8"))

    located = []
    if artist_id:
      albums = self.GetArtistAlbums(artist_id)
    else:
      albums = self.GetAlbums()
    if 'result' in albums and 'albums' in albums['result']:
      albums_list = albums['result']['albums']
      ll = self.matchHeard(heard_search, albums['result']['albums'])
      if ll:
        located = [(item['albumid'], item['label']) for item in ll]

    return located

  def FindSong(self, heard_search, artist_id=None, album_id=None):
    log.info('Searching for song "%s"', heard_search.encode("utf-8"))

    located = []
    if album_id:
      songs = self.GetAlbumSongs(album_id)
    elif artist_id:
      songs = self.GetArtistSongs(artist_id)
    else:
      songs = self.GetSongs()
    if 'result' in songs and 'songs' in songs['result']:
      ll = self.matchHeard(heard_search, songs['result']['songs'])
      if ll:
        located = [(item['songid'], item['label']) for item in ll]

    return located

  def FindAddon(self, heard_search):
    log.info('Searching for addon "%s"', heard_search.encode("utf-8"))

    located = []
    for content in ['video', 'audio', 'image', 'executable']:
      addons = self.GetAddons(content)
      if 'result' in addons and 'addons' in addons['result']:
        ll = self.matchHeard(heard_search, addons['result']['addons'], 'name')
        if ll:
          located = [(item['addonid'], item['name']) for item in ll]

    return located


  # Playlists

  def ClearAudioPlaylist(self):
    return self.SendCommand(RPCString("Playlist.Clear", {"playlistid": 0}))

  def AddSongToPlaylist(self, song_id):
    return self.SendCommand(RPCString("Playlist.Add", {"playlistid": 0, "item": {"songid": int(song_id)}}))

  def AddSongsToPlaylist(self, song_ids, shuffle=False):
    songs_array = []

    if shuffle:
      random.shuffle(song_ids)

    songs_array = [dict(songid=song_id) for song_id in song_ids[:self.playlist_limit]]

    # Segment the requests into chunks that Kodi will accept in a single call
    for a in [songs_array[x:x+2000] for x in range(0, len(songs_array), 2000)]:
      log.info('Adding %d items to the queue...', len(a))
      res = self.SendCommand(RPCString("Playlist.Add", {"playlistid": 0, "item": a}))

    return res

  def AddAlbumToPlaylist(self, album_id, shuffle=False):
    songs_result = self.GetAlbumSongs(album_id)
    songs = songs_result['result']['songs']
    songs_array = []
    for song in songs:
      songs_array.append(song['songid'])

    return self.AddSongsToPlaylist(songs_array, shuffle)

  def GetAudioPlaylistItems(self):
    return self.SendCommand(RPCString("Playlist.GetItems", {"playlistid": 0}))

  def StartAudioPlaylist(self, playlist_file=None):
    if playlist_file:
      # Note that subsequent shuffle commands won't work with this, as Kodi
      # considers a playlist to be a single item.
      #
      # Further, Kodi seems to completely ignore "options":{"shuffled":True} here
      return self.SendCommand(RPCString("Player.Open", {"item": {"file": playlist_file}}), False)
    else:
      return self.SendCommand(RPCString("Player.Open", {"item": {"playlistid": 0}}), False)

  def ClearVideoPlaylist(self):
    return self.SendCommand(RPCString("Playlist.Clear", {"playlistid": 1}))

  def AddEpisodeToPlayList(self, ep_id):
    return self.SendCommand(RPCString("Playlist.Add", {"playlistid": 1, "item": {"episodeid": int(ep_id)}}))

  def AddEpisodesToPlaylist(self, episode_ids, shuffle=False):
    if shuffle:
      random.shuffle(episode_ids)

    episodes_array = [dict(episodeid=episode_id) for episode_id in episode_ids[:self.playlist_limit]]

    # Segment the requests into chunks that Kodi will accept in a single call
    episode_groups = [episodes_array[x:x+2000] for x in range(0, len(episodes_array), 2000)]
    for a in episode_groups:
      log.info('Adding %d items to the queue...', len(a))
      res = self.SendCommand(RPCString("Playlist.Add", {"playlistid": 1, "item": a}))

    return res

  def AddMusicVideosToPlaylist(self, musicvideo_ids, shuffle=False):
    if shuffle:
      random.shuffle(musicvideo_ids)

    musicvideos_array = [dict(musicvideoid=musicvideo_id) for musicvideo_id in musicvideo_ids[:self.playlist_limit]]

    # Segment the requests into chunks that Kodi will accept in a single call
    musicvideo_groups = [musicvideos_array[x:x+2000] for x in range(0, len(musicvideos_array), 2000)]
    for a in musicvideo_groups:
      log.info('Adding %d items to the queue...', len(a))
      res = self.SendCommand(RPCString("Playlist.Add", {"playlistid": 1, "item": a}))

    return res

  def AddMovieToPlaylist(self, movie_id):
    return self.SendCommand(RPCString("Playlist.Add", {"playlistid": 1, "item": {"movieid": int(movie_id)}}))

  def AddVideosToPlaylist(self, video_files, shuffle=False):
    if shuffle:
      random.shuffle(video_files)

    videos_array = [dict(file=video_file) for video_file in video_files[:self.playlist_limit]]

    # Segment the requests into chunks that Kodi will accept in a single call
    video_groups = [videos_array[x:x+2000] for x in range(0, len(videos_array), 2000)]
    for a in video_groups:
      log.info('Adding %d items to the queue...', len(a))
      res = self.SendCommand(RPCString("Playlist.Add", {"playlistid": 1, "item": a}))

    return res

  def GetVideoPlaylistItems(self):
    return self.SendCommand(RPCString("Playlist.GetItems", {"playlistid": 1}))

  # Note that subsequent shuffle commands won't work with this, as Kodi
  # considers a playlist to be a single item.
  def StartVideoPlaylist(self, playlist_file=None):
    if playlist_file:
      return self.SendCommand(RPCString("Player.Open", {"item": {"file": playlist_file}}), False)
    else:
      return self.SendCommand(RPCString("Player.Open", {"item": {"playlistid": 1}}), False)


  # Direct plays

  def PlayFile(self, path):
    return self.SendCommand(RPCString("Player.Open", {"item": {"file": path}}), False)

  def PlayEpisode(self, ep_id, resume=True):
    return self.SendCommand(RPCString("Player.Open", {"item": {"episodeid": ep_id}, "options": {"resume": resume}}), False)

  def PlayMovie(self, movie_id, resume=True):
    return self.SendCommand(RPCString("Player.Open", {"item": {"movieid": movie_id}, "options": {"resume": resume}}), False)

  def PlayMusicVideo(self, musicvideo_id):
    return self.SendCommand(RPCString("Player.Open", {"item": {"musicvideoid": musicvideo_id}}), False)

  def PartyPlayMusic(self):
    return self.SendCommand(RPCString("Player.Open", {"item": {"partymode": "music"}}), False)


  # Tell Kodi to update its video or music libraries

  def UpdateVideo(self):
    self.cache.clear()
    return self.SendCommand(RPCString("VideoLibrary.Scan"), False)

  def CleanVideo(self):
    self.cache.clear()
    return self.SendCommand(RPCString("VideoLibrary.Clean"), False)

  def UpdateMusic(self):
    self.cache.clear()
    return self.SendCommand(RPCString("AudioLibrary.Scan"), False)

  def CleanMusic(self):
    self.cache.clear()
    return self.SendCommand(RPCString("AudioLibrary.Clean"), False)


  # Perform UI actions that match the normal remote control buttons

  def PageUp(self):
    return self.SendCommand(RPCString("Input.ExecuteAction", {"action": "pageup"}), False)

  def PageDown(self):
    return self.SendCommand(RPCString("Input.ExecuteAction", {"action": "pagedown"}), False)

  def ToggleWatched(self):
    return self.SendCommand(RPCString("Input.ExecuteAction", {"action": "togglewatched"}))

  def Info(self):
    return self.SendCommand(RPCString("Input.Info"), False)

  def Menu(self):
    return self.SendCommand(RPCString("Input.ContextMenu"), False)

  def Home(self):
    return self.SendCommand(RPCString("Input.Home"), False)

  def Select(self):
    return self.SendCommand(RPCString("Input.Select"), False)

  def Up(self):
    return self.SendCommand(RPCString("Input.Up"), False)

  def Down(self):
    return self.SendCommand(RPCString("Input.Down"), False)

  def Left(self):
    return self.SendCommand(RPCString("Input.Left"), False)

  def Right(self):
    return self.SendCommand(RPCString("Input.Right"), False)

  def Back(self):
    return self.SendCommand(RPCString("Input.Back"), False)

  def DownloadSubtitles(self):
    return self.SendCommand(RPCString("GUI.ActivateWindow", {"window": "subtitlesearch"}), False)

  def ShowMovies(self, genre_id=None):
    if genre_id:
      win = 'videodb://movies/genres/%d/' % (genre_id)
    else:
      win = 'MovieTitles'
    return self.SendCommand(RPCString("GUI.ActivateWindow", {"window": "videos", "parameters": [win, "return"]}), False)

  def ShowTvShows(self, genre_id=None):
    if genre_id:
      win = 'videodb://tvshows/genres/%d/' % (genre_id)
    else:
      win = 'TVShowTitles'
    return self.SendCommand(RPCString("GUI.ActivateWindow", {"window": "videos", "parameters": [win, "return"]}), False)

  def ShowMusicVideos(self, genre_id=None):
    if genre_id:
      win = 'videodb://musicvideos/genres/%d/' % (genre_id)
    else:
      win = 'MusicVideoTitles'
    return self.SendCommand(RPCString("GUI.ActivateWindow", {"window": "videos", "parameters": [win, "return"]}), False)

  def ShowMusic(self, genre_id=None):
    if genre_id:
      win = 'musicdb://genres/%d/' % (genre_id)
      return self.SendCommand(RPCString("GUI.ActivateWindow", {"window": "music", "parameters": [win, "return"]}), False)
    else:
      return self.SendCommand(RPCString("GUI.ActivateWindow", {"window": "music"}), False)

  def ShowMusicArtists(self):
    return self.SendCommand(RPCString("GUI.ActivateWindow", {"window": "music", "parameters": ["Artists", "return"]}), False)

  def ShowMusicAlbums(self):
    return self.SendCommand(RPCString("GUI.ActivateWindow", {"window": "music", "parameters": ["Albums", "return"]}), False)

  def ShowVideoPlaylist(self, playlist_path):
    return self.SendCommand(RPCString("GUI.ActivateWindow", {"window": "videos", "parameters": [playlist_path, "return"]}), False)

  def ShowMusicPlaylist(self, playlist_path):
    return self.SendCommand(RPCString("GUI.ActivateWindow", {"window": "music", "parameters": [playlist_path, "return"]}), False)

  def ToggleFullscreen(self):
    return self.SendCommand(RPCString("GUI.SetFullscreen", {"fullscreen": "toggle"}), False)

  def ToggleStereoscopicMode(self):
    return self.SendCommand(RPCString("Input.ExecuteAction", {"action": "togglestereomode"}))

  def ToggleAudioPassthrough(self):
    return self.SendCommand(RPCString("Input.ExecuteAction", {"action": "audiotoggledigital"}), False)

  def ToggleMute(self):
    return self.SendCommand(RPCString("Application.SetMute", {"mute": "toggle"}), False)

  def GetCurrentVolume(self):
    return self.SendCommand(RPCString("Application.GetProperties", fields=["volume", "muted"]))

  def VolumeUp(self):
    resp = self.GetCurrentVolume()
    vol = resp['result']['volume']
    if vol % 10 == 0:
      # already modulo 10, so just add 10
      vol += 10
    else:
      # round up to nearest 10
      vol -= vol % -10
    if vol > 100:
      vol = 100
    return self.SendCommand(RPCString("Application.SetVolume", {"volume": vol}))

  def VolumeDown(self):
    resp = self.GetCurrentVolume()
    vol = resp['result']['volume']
    if vol % 10 != 0:
      # round up to nearest 10 first
      vol -= vol % -10
    vol -= 10
    if vol < 0:
      vol = 0
    return self.SendCommand(RPCString("Application.SetVolume", {"volume": vol}))

  def VolumeSet(self, vol, percent=True):
    if vol < 0:
      vol = 0
    if not percent:
      # specified with scale of 0 to 10
      vol *= 10
    if vol > 100:
      vol = 100
    return self.SendCommand(RPCString("Application.SetVolume", {"volume": vol}))

  def SendText(self, send_text):
    return self.SendCommand(RPCString("Input.SendText", {"done": False, "text": send_text}))


  # Player controls

  def PlayerPlayPause(self):
    playerid = self.GetPlayerID()
    if playerid is not None:
      return self.SendCommand(RPCString("Player.PlayPause", {"playerid": playerid}), False)

  def PlayerSkip(self):
    playerid = self.GetPlayerID()
    if playerid is not None:
      return self.SendCommand(RPCString("Player.GoTo", {"playerid": playerid, "to": "next"}), False)

  def PlayerPrev(self):
    playerid = self.GetPlayerID()
    if playerid is not None:
      self.SendCommand(RPCString("Player.GoTo", {"playerid": playerid, "to": "previous"}))
      return self.SendCommand(RPCString("Player.GoTo", {"playerid": playerid, "to": "previous"}), False)

  def PlayerStartOver(self):
    playerid = self.GetPlayerID()
    if playerid is not None:
      return self.SendCommand(RPCString("Player.Seek", {"playerid": playerid, "value": 0}), False)

  def PlayerStop(self):
    playerid = self.GetPlayerID()
    if playerid is not None:
      return self.SendCommand(RPCString("Player.Stop", {"playerid": playerid}))

  def PlayerSeek(self, seconds):
    playerid = self.GetPlayerID()
    if playerid:
      return self.SendCommand(RPCString("Player.Seek", {"playerid": playerid, "value": {"seconds": seconds}}), False)

  def PlayerSeekSmallForward(self):
    playerid = self.GetPlayerID()
    if playerid:
      return self.SendCommand(RPCString("Player.Seek", {"playerid": playerid, "value": "smallforward"}), False)

  def PlayerSeekSmallBackward(self):
    playerid = self.GetPlayerID()
    if playerid:
      return self.SendCommand(RPCString("Player.Seek", {"playerid": playerid, "value": "smallbackward"}), False)

  def PlayerSeekBigForward(self):
    playerid = self.GetPlayerID()
    if playerid:
      return self.SendCommand(RPCString("Player.Seek", {"playerid": playerid, "value": "bigforward"}), False)

  def PlayerSeekBigBackward(self):
    playerid = self.GetPlayerID()
    if playerid:
      return self.SendCommand(RPCString("Player.Seek", {"playerid": playerid, "value": "bigbackward"}), False)

  def PlayerShuffleOn(self):
    playerid = self.GetPlayerID()
    if playerid is not None:
      return self.SendCommand(RPCString("Player.SetShuffle", {"playerid": playerid, "shuffle": True}))

  def PlayerShuffleOff(self):
    playerid = self.GetPlayerID()
    if playerid is not None:
      return self.SendCommand(RPCString("Player.SetShuffle", {"playerid": playerid, "shuffle": False}))

  def PlayerLoopOn(self):
    playerid = self.GetPlayerID()
    if playerid is not None:
      return self.SendCommand(RPCString("Player.SetRepeat", {"playerid": playerid, "repeat": "cycle"}))

  def PlayerLoopOff(self):
    playerid = self.GetPlayerID()
    if playerid is not None:
      return self.SendCommand(RPCString("Player.SetRepeat", {"playerid": playerid, "repeat": "off"}))

  def PlayerSubtitlesOn(self):
    playerid = self.GetVideoPlayerID()
    if playerid:
      return self.SendCommand(RPCString("Player.SetSubtitle", {"playerid": playerid, "subtitle": "on"}))

  def PlayerSubtitlesOff(self):
    playerid = self.GetVideoPlayerID()
    if playerid:
      return self.SendCommand(RPCString("Player.SetSubtitle", {"playerid": playerid, "subtitle": "off"}))

  def PlayerSubtitlesNext(self):
    playerid = self.GetVideoPlayerID()
    if playerid:
      return self.SendCommand(RPCString("Player.SetSubtitle", {"playerid": playerid, "subtitle": "next", "enable": True}))

  def PlayerSubtitlesPrevious(self):
    playerid = self.GetVideoPlayerID()
    if playerid:
      return self.SendCommand(RPCString("Player.SetSubtitle", {"playerid": playerid, "subtitle": "previous", "enable": True}))

  def PlayerAudioStreamNext(self):
    playerid = self.GetVideoPlayerID()
    if playerid:
      return self.SendCommand(RPCString("Player.SetAudioStream", {"playerid": playerid, "stream": "next"}))

  def PlayerAudioStreamPrevious(self):
    playerid = self.GetVideoPlayerID()
    if playerid:
      return self.SendCommand(RPCString("Player.SetAudioStream", {"playerid": playerid, "stream": "previous"}))

  def PlayerMoveUp(self):
    playerid = self.GetPicturePlayerID()
    if playerid:
      return self.SendCommand(RPCString("Player.Move", {"playerid": playerid, "direction": "up"}), False)

  def PlayerMoveDown(self):
    playerid = self.GetPicturePlayerID()
    if playerid:
      return self.SendCommand(RPCString("Player.Move", {"playerid": playerid, "direction": "down"}), False)

  def PlayerMoveLeft(self):
    playerid = self.GetPicturePlayerID()
    if playerid:
      return self.SendCommand(RPCString("Player.Move", {"playerid": playerid, "direction": "left"}), False)

  def PlayerMoveRight(self):
    playerid = self.GetPicturePlayerID()
    if playerid:
      return self.SendCommand(RPCString("Player.Move", {"playerid": playerid, "direction": "right"}), False)

  def PlayerZoom(self, lvl=0):
    playerid = self.GetPicturePlayerID()
    if playerid and lvl > 0 and lvl < 11:
      return self.SendCommand(RPCString("Player.Zoom", {"playerid": playerid, "zoom": lvl}), False)

  def PlayerZoomIn(self):
    playerid = self.GetPicturePlayerID()
    if playerid:
      return self.SendCommand(RPCString("Player.Zoom", {"playerid": playerid, "zoom": "in"}), False)

  def PlayerZoomOut(self):
    playerid = self.GetPicturePlayerID()
    if playerid:
      return self.SendCommand(RPCString("Player.Zoom", {"playerid": playerid, "zoom": "out"}), False)

  def PlayerRotateClockwise(self):
    playerid = self.GetPicturePlayerID()
    if playerid:
      return self.SendCommand(RPCString("Player.Rotate", {"playerid": playerid, "value": "clockwise"}), False)

  def PlayerRotateCounterClockwise(self):
    playerid = self.GetPicturePlayerID()
    if playerid:
      return self.SendCommand(RPCString("Player.Rotate", {"playerid": playerid, "value": "counterclockwise"}), False)


  # Addons

  def AddonExecute(self, addon_id, params={}):
    return self.SendCommand(RPCString("Addons.ExecuteAddon", {"addonid": addon_id, "params": params}))

  def AddonGlobalSearch(self, needle=''):
    return self.AddonExecute("script.globalsearch", {"searchstring": needle.encode("utf-8")})

  def AddonCinemaVision(self):
    return self.AddonExecute("script.cinemavision", ["experience"])


  # Library queries

  # mediatype should be one of:
  #   movies, tvshows, episodes, musicvideos, artists, albums, songs
  #
  # returns a list like:
  #   [type, label, library_id, genre]
  #
  # where type is one of:
  #   movie, tvshow, episode, musicvideo, artist, album, song
  def GetRecommendedItem(self, mediatype=None, mediagenre=None):
    answer = ['', '', 0, mediagenre]

    if not mediatype:
      return answer

    m = []
    if mediatype == 'movies':
      if mediagenre:
        m = self.GetUnwatchedMoviesByGenre(mediagenre, sort=SORT_RATING, limits=LIMIT_RECOMMENDED_MOVIES)
      else:
        m = self.GetUnwatchedMovies(sort=SORT_RATING, limits=LIMIT_RECOMMENDED_MOVIES)
      if not m:
        # Fall back to all movies if no unwatched available
        if mediagenre:
          movies = self.GetMoviesByGenre(mediagenre, sort=SORT_RATING, limits=LIMIT_RECOMMENDED_MOVIES)
        else:
          movies = self.GetMovies(sort=SORT_RATING, limits=LIMIT_RECOMMENDED_MOVIES)
        if 'result' in movies and 'movies' in movies['result']:
          m = movies['result']['movies']
      if m:
        r = random.choice(m)
        answer[0] = 'movie'
        answer[1] = r['label']
        answer[2] = r['movieid']
    elif mediatype == 'tvshows':
      if mediagenre:
        m = self.GetUnwatchedShowsByGenre(mediagenre, sort=SORT_RATING, limits=LIMIT_RECOMMENDED_SHOWS)
      else:
        m = self.GetUnwatchedShows(sort=SORT_RATING, limits=LIMIT_RECOMMENDED_SHOWS)
      if not m:
        # Fall back to all shows if no unwatched available
        if mediagenre:
          shows = self.GetShowsByGenre(mediagenre, sort=SORT_RATING, limits=LIMIT_RECOMMENDED_SHOWS)
        else:
          shows = self.GetShows(sort=SORT_RATING, limits=LIMIT_RECOMMENDED_SHOWS)
        if 'result' in shows and 'tvshows' in shows['result']:
          m = shows['result']['tvshows']
      if m:
        r = random.choice(m)
        answer[0] = 'tvshow'
        answer[1] = r['label']
        answer[2] = r['tvshowid']
    elif mediatype == 'episodes':
      if mediagenre:
        shows = self.GetUnwatchedShowsByGenre(mediagenre, sort=SORT_RATING, limits=LIMIT_RECOMMENDED_SHOWS)
      else:
        shows = self.GetUnwatchedShows(sort=SORT_RATING, limits=LIMIT_RECOMMENDED_SHOWS)
      if shows:
        r = random.choice(shows)
        m = self.GetUnwatchedEpisodesFromShow(r['tvshowid'], limits=(0, 1))
      if not m:
        # Fall back to all episodes if no unwatched available
        if mediagenre:
          episodes = self.GetEpisodesByGenre(mediagenre, sort=SORT_RATING, limits=LIMIT_RECOMMENDED_EPISODES)
        else:
          episodes = self.GetEpisodes(sort=SORT_RATING, limits=LIMIT_RECOMMENDED_EPISODES)
        if 'result' in episodes and 'episodes' in episodes['result']:
          m = episodes['result']['episodes']
      if m:
        r = random.choice(m)
        answer[0] = 'episode'
        answer[1] = r['label']
        answer[2] = r['episodeid']
    elif mediatype == 'musicvideos':
      if mediagenre:
        musicvideos = self.GetMusicVideosByGenre(mediagenre, sort=SORT_RATING, limits=LIMIT_RECOMMENDED_MUSICVIDEOS)
      else:
        musicvideos = self.GetMusicVideos(sort=SORT_RATING, limits=LIMIT_RECOMMENDED_MUSICVIDEOS)
      if 'result' in musicvideos and 'musicvideos' in musicvideos['result']:
        m = musicvideos['result']['musicvideos']
      if m:
        r = random.choice(m)
        answer[0] = 'musicvideo'
        answer[1] = r['label']
        answer[2] = r['musicvideoid']
    elif mediatype == 'artists':
      if mediagenre:
        artists = self.GetMusicArtistsByGenre(mediagenre, sort=SORT_RATING, limits=LIMIT_RECOMMENDED_ARTISTS)
      else:
        artists = self.GetMusicArtists(sort=SORT_RATING, limits=LIMIT_RECOMMENDED_ARTISTS)
      if 'result' in artists and 'artists' in artists['result']:
        m = artists['result']['artists']
      if m:
        r = random.choice(m)
        answer[0] = 'artist'
        answer[1] = r['label']
        answer[2] = r['artistid']
    elif mediatype == 'albums':
      if mediagenre:
        albums = self.GetAlbumsByGenre(mediagenre, sort=SORT_RATING, limits=LIMIT_RECOMMENDED_ALBUMS)
      else:
        albums = self.GetAlbums(sort=SORT_RATING, limits=LIMIT_RECOMMENDED_ALBUMS)
      if 'result' in albums and 'albums' in albums['result']:
        m = albums['result']['albums']
      if m:
        r = random.choice(m)
        answer[0] = 'album'
        answer[1] = r['label']
        answer[2] = r['albumid']
    elif mediatype == 'songs':
      if mediagenre:
        songs = self.GetSongsByGenre(mediagenre, sort=SORT_RATING, limits=LIMIT_RECOMMENDED_SONGS)
      else:
        songs = self.GetSongs(sort=SORT_RATING, limits=LIMIT_RECOMMENDED_SONGS)
      if 'result' in songs and 'songs' in songs['result']:
        m = songs['result']['songs']
      if m:
        r = random.choice(m)
        answer[0] = 'song'
        answer[1] = r['label']
        answer[2] = r['songid']

    return answer

  def GetRecommendedVideoItem(self):
    answer = []
    items = []
    for content in ['movies', 'tvshows', 'episodes', 'musicvideos']:
      item = self.GetRecommendedItem(content)
      if item[0]:
        items.append(item)

    if items:
      answer = random.choice(items)

    return answer

  def GetRecommendedAudioItem(self):
    answer = []
    items = []
    for content in ['musicvideos', 'artists', 'albums', 'songs']:
      item = self.GetRecommendedItem(content)
      if item[0]:
        items.append(item)

    if items:
      answer = random.choice(items)

    return answer

  # content can be: video, audio, image, executable, or unknown
  def GetAddons(self, content):
    if content:
      return self.SendCommand(RPCString("Addons.GetAddons", {"content": content}, fields=["name"]))
    else:
      return self.SendCommand(RPCString("Addons.GetAddons", fields=["name"]))

  def GetAddonDetails(self, addon_id):
    return self.SendCommand(RPCString("Addons.GetAddonDetails", {"addonid": addon_id}, fields=["name", "version", "description", "summary"]))

  def GetPlaylistItems(self, playlist_file):
    return self.SendCommand(RPCString("Files.GetDirectory", {"directory": playlist_file}))

  def GetMusicPlaylists(self):
    return self.SendCommand(RPCString("Files.GetDirectory", {"directory": "special://musicplaylists"}))

  def GetMusicArtists(self, sort=None, filters=None, filtertype=None, limits=None):
    return self.SendCommand(RPCString("AudioLibrary.GetArtists", {"albumartistsonly": False}, sort=sort, filters=filters, filtertype=filtertype, limits=limits), cache_resp=True)

  def GetMusicArtistsByGenre(self, genre, sort=None, limits=None):
    return self.GetMusicArtists(sort=sort, filters=[{"field": "genre", "operator": "is", "value": genre}], limits=limits)

  def GetMusicGenres(self):
    return self.SendCommand(RPCString("AudioLibrary.GetGenres"), cache_resp=True)

  def GetArtistAlbums(self, artist_id):
    return self.SendCommand(RPCString("AudioLibrary.GetAlbums", filters=[{"artistid": int(artist_id)}]), cache_resp=True)

  def GetNewestAlbumFromArtist(self, artist_id):
    data = self.SendCommand(RPCString("AudioLibrary.GetAlbums", sort=SORT_YEAR, filters=[{"artistid": int(artist_id)}], limits=(0, 1)), cache_resp=True)
    if 'albums' in data['result']:
      album = data['result']['albums'][0]
      return album['albumid']
    else:
      return None

  def GetSongs(self, sort=None, filters=None, filtertype=None, limits=None):
    return self.SendCommand(RPCString("AudioLibrary.GetSongs", sort=sort, filters=filters, filtertype=filtertype, limits=limits), cache_resp=True)

  def GetSongsByGenre(self, genre, sort=None, limits=None):
    return self.GetSongs(sort=sort, filters=[{"field": "genre", "operator": "is", "value": genre}], limits=limits)

  def GetSongsPath(self):
    return self.SendCommand(RPCString("AudioLibrary.GetSongs", fields=["file"]))

  def GetSongIdPath(self, song_id):
    return self.SendCommand(RPCString("AudioLibrary.GetSongDetails", {"songid": int(song_id)}, fields=["file"]))

  def GetSongDetails(self, song_id):
    data = self.SendCommand(RPCString("AudioLibrary.GetSongDetails", {"songid": int(song_id)}, fields=["artist"]))
    return data['result']['songdetails']

  def GetArtistSongs(self, artist_id, sort=None, limits=None):
    return self.GetSongs(sort=sort, filters=[{"artistid": int(artist_id)}], limits=limits)

  def GetArtistSongsByGenre(self, artist, genre, sort=None, limits=None):
    return self.GetSongs(sort=sort, filters=[{"field": "artist", "operator": "is", "value": artist}, {"field": "genre", "operator": "is", "value": genre}], limits=limits)

  def GetArtistSongsPath(self, artist_id):
    return self.SendCommand(RPCString("AudioLibrary.GetSongs", filters=[{"artistid": int(artist_id)}], fields=["file"]), cache_resp=True)

  def GetAlbums(self, sort=None, filters=None, filtertype=None, limits=None):
    return self.SendCommand(RPCString("AudioLibrary.GetAlbums", sort=sort, filters=filters, filtertype=filtertype, limits=limits), cache_resp=True)

  def GetAlbumsByGenre(self, genre, sort=None, limits=None):
    return self.GetAlbums(sort=sort, filters=[{"field": "genre", "operator": "is", "value": genre}], limits=limits)

  def GetAlbumDetails(self, album_id):
    data = self.SendCommand(RPCString("AudioLibrary.GetAlbumDetails", {"albumid": int(album_id)}, fields=["artist"]))
    return data['result']['albumdetails']

  def GetAlbumSongs(self, album_id, sort=None, limits=None):
    return self.GetSongs(sort=sort, filters=[{"albumid": int(album_id)}], limits=limits)

  def GetAlbumSongsPath(self, album_id):
    return self.SendCommand(RPCString("AudioLibrary.GetSongs", filters=[{"albumid": int(album_id)}], fields=["file"]), cache_resp=True)

  def GetRecentlyAddedAlbums(self):
    return self.SendCommand(RPCString("AudioLibrary.GetRecentlyAddedAlbums", fields=["artist"]))

  def GetRecentlyAddedSongs(self):
    return self.SendCommand(RPCString("AudioLibrary.GetRecentlyAddedSongs", fields=["artist"]))

  def GetRecentlyAddedSongsPath(self):
    return self.SendCommand(RPCString("AudioLibrary.GetRecentlyAddedSongs", fields=["artist", "file"]))

  def GetVideoPlaylists(self):
    return self.SendCommand(RPCString("Files.GetDirectory", {"directory": "special://videoplaylists"}))

  def GetVideoGenres(self, genretype='movie'):
    return self.SendCommand(RPCString("VideoLibrary.GetGenres", {"type": genretype}), cache_resp=True)

  def GetMusicVideos(self, sort=None, filters=None, filtertype=None, limits=None):
    return self.SendCommand(RPCString("VideoLibrary.GetMusicVideos", fields=["artist"], sort=sort, filters=filters, filtertype=filtertype, limits=limits), cache_resp=True)

  def GetMusicVideosByGenre(self, genre, sort=None, limits=None):
    return self.GetMusicVideos(sort=sort, filters=[{"genre": genre}], limits=None)

  def GetMusicVideoDetails(self, mv_id):
    data = self.SendCommand(RPCString("VideoLibrary.GetMusicVideoDetails", {"musicvideoid": int(mv_id)}, fields=["artist"]))
    return data['result']['musicvideodetails']

  def GetMovies(self, sort=None, filters=None, filtertype=None, limits=None):
    return self.SendCommand(RPCString("VideoLibrary.GetMovies", sort=sort, filters=filters, filtertype=filtertype, limits=limits), cache_resp=True)

  def GetMoviesByGenre(self, genre, sort=None, limits=None):
    return self.GetMovies(sort=sort, fiters=[{"genre": genre}], limits=limits)

  def GetMovieDetails(self, movie_id):
    data = self.SendCommand(RPCString("VideoLibrary.GetMovieDetails", {"movieid": movie_id}, fields=["resume", "trailer"]))
    return data['result']['moviedetails']

  def GetShows(self, sort=None, filters=None, filtertype=None, limits=None):
    return self.SendCommand(RPCString("VideoLibrary.GetTVShows", sort=sort, filters=filters, filtertype=filtertype, limits=limits), cache_resp=True)

  def GetShowsByGenre(self, genre, sort=None, limits=None):
    return self.GetShows(sort=sort, filters=[{"genre": genre}], limits=limits)

  def GetShowDetails(self, show_id):
    data = self.SendCommand(RPCString("VideoLibrary.GetTVShowDetails", {"tvshowid": show_id}, fields=["art"]))
    return data['result']['tvshowdetails']

  def GetEpisodes(self, sort=None, filters=None, filtertype=None, limits=None):
    return self.SendCommand(RPCString("VideoLibrary.GetEpisodes", sort=sort, filters=filters, filtertype=filtertype, limits=limits), cache_resp=True)

  def GetEpisodesByGenre(self, genre, sort=None, limits=None):
    return self.GetEpisodes(sort=sort, filters=[{"field": "genre", "operator": "is", "value": genre}], limits=limits)

  def GetEpisodesFromShow(self, show_id):
    return self.SendCommand(RPCString("VideoLibrary.GetEpisodes", {"tvshowid": int(show_id)}), cache_resp=True)

  def GetEpisodeDetails(self, ep_id):
    data = self.SendCommand(RPCString("VideoLibrary.GetEpisodeDetails", {"episodeid": int(ep_id)}, fields=["showtitle", "season", "episode", "resume"]))
    return data['result']['episodedetails']

  def GetNewestEpisodeFromShow(self, show_id):
    data = self.SendCommand(RPCString("VideoLibrary.GetEpisodes", {"tvshowid": int(show_id)}, sort=SORT_DATEADDED, limits=(0, 1)))
    if 'episodes' in data['result']:
      episode = data['result']['episodes'][0]
      return episode['episodeid']
    else:
      return None

  def GetNextUnwatchedEpisode(self, show_id):
    data = self.SendCommand(RPCString("VideoLibrary.GetEpisodes", {"tvshowid": int(show_id)}, filters=[FILTER_UNWATCHED], sort=SORT_EPISODE, fields=["playcount"], limits=(0, 1)))
    if 'episodes' in data['result']:
      episode = data['result']['episodes'][0]
      return episode['episodeid']
    else:
      return None

  def GetLastWatchedShow(self):
    return self.SendCommand(RPCString("VideoLibrary.GetEpisodes", sort=SORT_LASTPLAYED, filters=[FILTER_WATCHED, {"field": "lastplayed", "operator": "isnot", "value": "0"}], fields=["tvshowid", "showtitle"], limits=(0, 1)))

  def GetSpecificEpisode(self, show_id, season, episode):
    data = self.SendCommand(RPCString("VideoLibrary.GetEpisodes", {"tvshowid": int(show_id), "season": int(season)}, fields=["season", "episode"]))
    if 'episodes' in data['result']:
      correct_id = None
      for episode_data in data['result']['episodes']:
        if int(episode_data['episode']) == int(episode):
          correct_id = episode_data['episodeid']
          break

      return correct_id
    else:
      return None

  def GetEpisodesFromShowDetails(self, show_id):
    return self.SendCommand(RPCString("VideoLibrary.GetEpisodes", {"tvshowid": int(show_id)}, fields=["season", "episode"]))

  # Returns a list of dictionaries with information about unwatched movies. Useful for
  # telling/showing users what's ready to be watched. Setting max to very high values
  # can take a long time.
  def GetUnwatchedMovies(self, sort=SORT_DATEADDED, limits=None):
    if not limits:
      limits = (0, self.max_unwatched_movies)
    data = self.SendCommand(RPCString("VideoLibrary.GetMovies", sort=sort, filters=[FILTER_UNWATCHED], fields=["title", "playcount", "dateadded"], limits=limits))
    answer = []
    if 'movies' in data['result']:
      for d in data['result']['movies']:
        answer.append({'title': d['title'], 'movieid': d['movieid'], 'label': d['label'], 'dateadded': datetime.datetime.strptime(d['dateadded'], "%Y-%m-%d %H:%M:%S")})
    return answer

  # Returns a list of dictionaries with information about unwatched movies in a particular genre. Useful for
  # telling/showing users what's ready to be watched. Setting max to very high values
  # can take a long time.
  def GetUnwatchedMoviesByGenre(self, genre, sort=SORT_DATEADDED, limits=None):
    if not limits:
      limits = (0, self.max_unwatched_movies)
    data = self.SendCommand(RPCString("VideoLibrary.GetMovies", sort=sort, filters=[FILTER_UNWATCHED, {"field": "genre", "operator": "contains", "value": genre}], fields=["title", "playcount", "dateadded"], limits=limits))
    answer = []
    if 'movies' in data['result']:
      for d in data['result']['movies']:
        answer.append({'title': d['title'], 'movieid': d['movieid'], 'label': d['label'], 'dateadded': datetime.datetime.strptime(d['dateadded'], "%Y-%m-%d %H:%M:%S")})
    return answer

  # Returns a list of dictionaries with information about unwatched shows. Useful for
  # telling/showing users what's ready to be watched. Setting max to very high values
  # can take a long time.
  def GetUnwatchedShows(self, sort=SORT_DATEADDED, limits=None):
    if not limits:
      limits = (0, self.max_unwatched_shows)
    data = self.SendCommand(RPCString("VideoLibrary.GetTVShows", sort=sort, filters=[FILTER_UNWATCHED], fields=["title", "playcount", "dateadded"], limits=limits))
    answer = []
    if 'tvshows' in data['result']:
      for d in data['result']['tvshows']:
        answer.append({'title': d['title'], 'tvshowid': d['tvshowid'], 'label': d['label'], 'dateadded': datetime.datetime.strptime(d['dateadded'], "%Y-%m-%d %H:%M:%S")})
    return answer

  # Returns a list of dictionaries with information about unwatched shows in a particular genre. Useful for
  # telling/showing users what's ready to be watched. Setting max to very high values
  # can take a long time.
  def GetUnwatchedShowsByGenre(self, genre, sort=SORT_DATEADDED, limits=None):
    if not limits:
      limits = (0, self.max_unwatched_shows)
    data = self.SendCommand(RPCString("VideoLibrary.GetTVShows", sort=sort, filters=[FILTER_UNWATCHED, {"field": "genre", "operator": "contains", "value": genre}], fields=["title", "playcount", "dateadded"], limits=limits))
    answer = []
    if 'tvshows' in data['result']:
      for d in data['result']['tvshows']:
        answer.append({'title': d['title'], 'tvshowid': d['tvshowid'], 'label': d['label'], 'dateadded': datetime.datetime.strptime(d['dateadded'], "%Y-%m-%d %H:%M:%S")})
    return answer

  # Returns a list of dictionaries with information about episodes that have been watched.
  def GetWatchedEpisodes(self, sort=None, limits=None):
    return self.SendCommand(RPCString("VideoLibrary.GetEpisodes", sort=sort, filters=[FILTER_WATCHED], fields=["playcount", "showtitle", "season", "episode", "lastplayed"], limits=limits))

  # Returns a list of dictionaries with information about unwatched episodes. Useful for
  # telling/showing users what's ready to be watched. Setting max to very high values
  # can take a long time.
  def GetUnwatchedEpisodes(self, sort=SORT_DATEADDED, limits=None):
    if not limits:
      limits = (0, self.max_unwatched_shows)
    data = self.SendCommand(RPCString("VideoLibrary.GetEpisodes", sort=sort, filters=[FILTER_UNWATCHED], fields=["title", "playcount", "showtitle", "tvshowid", "dateadded"], limits=limits))
    answer = []
    if 'episodes' in data['result']:
      shows = set([d['tvshowid'] for d in data['result']['episodes']])
      show_info = {}
      for show in shows:
        show_info[show] = self.GetShowDetails(show_id=show)
      for d in data['result']['episodes']:
        showinfo = show_info[d['tvshowid']]
        answer.append({'title': d['title'], 'episodeid': d['episodeid'], 'show': d['showtitle'], 'label': d['label'], 'dateadded': datetime.datetime.strptime(d['dateadded'], "%Y-%m-%d %H:%M:%S")})
    return answer

  def GetUnwatchedEpisodesFromShow(self, show_id, limits=None):
    data = self.SendCommand(RPCString("VideoLibrary.GetEpisodes", {"tvshowid": int(show_id)}, filters=[FILTER_UNWATCHED], fields=["title", "playcount", "showtitle", "tvshowid", "dateadded"], limits=limits))
    answer = []
    if 'episodes' in data['result']:
      for d in data['result']['episodes']:
        answer.append({'title': d['title'], 'episodeid': d['episodeid'], 'show': d['showtitle'], 'label': d['label'], 'dateadded': datetime.datetime.strptime(d['dateadded'], "%Y-%m-%d %H:%M:%S")})
    return answer


  # System commands

  def ApplicationQuit(self):
    return self.SendCommand(RPCString("Application.Quit"))

  def SystemHibernate(self):
    return self.SendCommand(RPCString("System.Hibernate"))

  def SystemReboot(self):
    return self.SendCommand(RPCString("System.Reboot"))

  def SystemShutdown(self):
    return self.SendCommand(RPCString("System.Shutdown"))

  def SystemSuspend(self):
    return self.SendCommand(RPCString("System.Suspend"))

  def SystemEjectMedia(self):
    return self.SendCommand(RPCString("System.EjectOpticalDrive"))


  # Misc helpers

  # Prepare file url for streaming
  def PrepareDownload(self, path=""):
    path = urllib.quote(path.encode('utf-8')).decode('utf-8')

    # Join the environment variables into a url
    url = "%s://%s:%s@%s:%s/%s/vfs" % (self.scheme, self.username, self.password, self.address, self.port, self.subpath)

    # Remove any double slashes in the url
    url = http_normalize_slashes(url)
    url = url + '/' + path

    accepted_answers = ['y', 'yes', 'Y', 'Yes', 'YES', 'true', 'True']

    if self.config.get(self.dev_cfg_section, 'use_proxy') in accepted_answers:
      stream_url = 'https://proxy.lexigr.am/proxy?file=' + url
    elif self.config.get(self.dev_cfg_section, 'alt_proxy'):
      stream_url = self.config.get(self.dev_cfg_section, 'alt_proxy') + url
    else:
      stream_url = url

    return stream_url

  # Get the first active player.
  def GetPlayerID(self, playertype=['picture', 'audio', 'video']):
    data = self.SendCommand(RPCString("Player.GetActivePlayers"))
    result = data.get("result", [])
    if result:
      for curitem in result:
        if curitem.get("type") in playertype:
          return curitem.get("playerid")
    return None

  # Get the first active Video player.
  def GetVideoPlayerID(self, playertype=['video']):
    data = self.SendCommand(RPCString("Player.GetActivePlayers"))
    result = data.get("result", [])
    if result:
      for curitem in result:
        if curitem.get("type") in playertype:
          return curitem.get("playerid")
    return None


  # Get the first active Audio player.
  def GetAudioPlayerID(self, playertype=['audio']):
    data = self.SendCommand(RPCString("Player.GetActivePlayers"))
    result = data.get("result", [])
    if result:
      for curitem in result:
        if curitem.get("type") in playertype:
          return curitem.get("playerid")
    return None

  # Get the first active Picture player.
  def GetPicturePlayerID(self, playertype=['picture']):
    data = self.SendCommand(RPCString("Player.GetActivePlayers"))
    result = data.get("result", [])
    if result:
      for curitem in result:
        if curitem.get("type") in playertype:
          return curitem.get("playerid")
    return None

  # Information about the video or audio that's currently playing

  def GetActivePlayItem(self):
    playerid = self.GetPlayerID()
    if playerid is not None:
      data = self.SendCommand(RPCString("Player.GetItem", {"playerid": playerid}, fields=["title", "album", "artist", "season", "episode", "showtitle", "tvshowid", "description"]))
      return data['result']['item']

  def GetActivePlayProperties(self):
    playerid = self.GetPlayerID()
    if playerid is not None:
      data = self.SendCommand(RPCString("Player.GetProperties", {"playerid": playerid}, fields=["currentaudiostream", "currentsubtitle", "canshuffle", "shuffled", "canrepeat", "repeat", "canzoom", "canrotate", "canmove"]))
      return data['result']

  # Returns current subtitles as a speakable string
  def GetCurrentSubtitles(self):
    subs = ""
    country_dic = getisocodes_dict()
    curprops = self.GetActivePlayProperties()
    if curprops is not None:
      try:
        # gets 3 character country code e.g. fre
        lang = curprops['currentsubtitle']['language']
        # looks up 3 character code in the dictionary e.g. fre|fra|fr|French|francais
        subslang = country_dic[lang]
        # matches 3 character code with the lang name
        subs = subslang[self.language]
        # joins full language name with the name of the subtitle file e.g. French External
        name = curprops['currentsubtitle']['name']
        if name:
          subs += " " + name
      except:
        pass
    return subs

  # Returns current audio stream as a speakable string
  def GetCurrentAudioStream(self):
    stream = ""
    country_dic = getisocodes_dict()
    curprops = self.GetActivePlayProperties()
    if curprops is not None:
      try:
        # gets 3 character country code e.g. fre
        lang = curprops['currentaudiostream']['language']
        # looks up 3 character code in the dictionary e.g. fre|fra|fr|French|francais
        streamlang = country_dic[lang]
        # matches 3 character code with the lang name
        stream = streamlang[self.language]
        # joins full language name with the name of the subtitle file e.g. French External
        name = curprops['currentaudiostream']['name']
        if name:
          stream += " " + name
      except:
        pass
    return stream

  # Returns information useful for building a progress bar to show an item's play time
  def GetPlayerStatus(self):
    playerid = self.GetVideoPlayerID()
    if playerid is None:
      playerid = self.GetAudioPlayerID()
    if playerid is not None:
      data = self.SendCommand(RPCString("Player.GetProperties", {"playerid": playerid}, fields=["percentage", "speed", "time", "totaltime"]))
      if 'result' in data:
        hours_total = data['result']['totaltime']['hours']
        hours_cur = data['result']['time']['hours']
        mins_total = hours_total * 60 + data['result']['totaltime']['minutes']
        mins_cur = hours_cur * 60 + data['result']['time']['minutes']
        speed = data['result']['speed']
        if hours_total > 0:
          total = '%d:%02d:%02d' % (hours_total, data['result']['totaltime']['minutes'], data['result']['totaltime']['seconds'])
          cur = '%d:%02d:%02d' % (data['result']['time']['hours'], data['result']['time']['minutes'], data['result']['time']['seconds'])
        else:
          total = '%02d:%02d' % (data['result']['totaltime']['minutes'], data['result']['totaltime']['seconds'])
          cur = '%02d:%02d' % (data['result']['time']['minutes'], data['result']['time']['seconds'])
        return {'state': 'play' if speed > 0 else 'pause', 'time': cur, 'time_hours': hours_cur, 'time_mins': mins_cur, 'totaltime': total, 'total_hours': hours_total, 'total_mins': mins_total, 'pct': data['result']['percentage']}
    return {'state': 'stop'}
