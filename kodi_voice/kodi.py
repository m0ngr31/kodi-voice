#!/usr/bin/env python

# For a complete discussion, see http://forum.kodi.tv/showthread.php?tid=254502

import datetime
import json
import requests
import time
import codecs
import urllib
import os
import random
import re
import string
import sys
import unicodedata
import roman
from num2words import num2words
from fuzzywuzzy import fuzz, process
from ConfigParser import SafeConfigParser


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


def RPCString(method, params=None):
  j = {"jsonrpc":"2.0", "method":method, "id":1}
  if params:
    j["params"] = params
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
      SHUTDOWN_MEANS_QUIT = os.getenv('SHUTDOWN_MEANS_QUIT')
      if SHUTDOWN_MEANS_QUIT and SHUTDOWN_MEANS_QUIT != 'None':
        self.set('DEFAULT', 'shutdown', SHUTDOWN_MEANS_QUIT)
      SKILL_TZ = os.getenv('SKILL_TZ')
      if SKILL_TZ and SKILL_TZ != 'None':
        self.set('DEFAULT', 'timezone', SKILL_TZ)
      LANGUAGE = os.getenv('LANGUAGE')
      if LANGUAGE and LANGUAGE != 'None':
        self.set('global', 'language', LANGUAGE)
      DEEP_SEARCH = os.getenv('DEEP_SEARCH')
      if DEEP_SEARCH and DEEP_SEARCH != 'None':
        self.set('global', 'deep_search', DEEP_SEARCH)
      MAX_UNWATCHED_EPISODES = os.getenv('MAX_UNWATCHED_EPISODES')
      if MAX_UNWATCHED_EPISODES and MAX_UNWATCHED_EPISODES != 'None':
        self.set('global', 'unwatched_episodes_max_results', MAX_UNWATCHED_EPISODES)
      MAX_UNWATCHED_MOVIES = os.getenv('MAX_UNWATCHED_MOVIES')
      if MAX_UNWATCHED_MOVIES and MAX_UNWATCHED_MOVIES != 'None':
        self.set('global', 'unwatched_movies_max_results', MAX_UNWATCHED_MOVIES)
      SKILL_APPID = os.getenv('SKILL_APPID')
      if SKILL_APPID and SKILL_APPID != 'None':
        self.set('alexa', 'skill_id', SKILL_APPID)
    else:
      self.config_file = config_file
      self.read(self.config_file)


class Kodi:
  def __init__(self, config=None, context=None):
    self.config = config

    # When testing from the web simulator there is no context object (04/2017)
    if context:
      self.deviceId = context.System.device.deviceId
    else:
      self.deviceId = 'Unknown Device'

    if config.has_section(self.deviceId):
      self.dev_cfg_section = self.deviceId
    else:
      self.dev_cfg_section = 'DEFAULT'

    self.language = self.config.get('global', 'language').lower()
    self.max_unwatched_episodes = int(self.config.get('global', 'unwatched_episodes_max_results'))
    self.max_unwatched_movies = int(self.config.get('global', 'unwatched_movies_max_results'))

    self.scheme   = self.config.get(self.dev_cfg_section, 'scheme')
    self.subpath  = self.config.get(self.dev_cfg_section, 'subpath')
    self.address  = self.config.get(self.dev_cfg_section, 'address')
    self.port     = self.config.get(self.dev_cfg_section, 'port')
    self.username = self.config.get(self.dev_cfg_section, 'username')
    self.password = self.config.get(self.dev_cfg_section, 'password')


  # Construct the JSON-RPC message and send it to the Kodi player
  def SendCommand(self, command):
    # Join the configuration variables into a url
    url = "%s://%s:%s/%s/%s" % (self.scheme, self.address, self.port, self.subpath, 'jsonrpc')

    # Remove any double slashes in the url
    url = http_normalize_slashes(url)

    print "Sending request to %s from device %s" % (url, self.deviceId)

    r = requests.post(url, data=command, auth=(self.username, self.password))
    if r.encoding is None:
      r.encoding = 'utf-8'

    try:
      return r.json()
    except:
      print "Error: json decoding failed {}".format(r)
      raise

  # Utilities

  def sanitize_name(self, *args, **kwargs):
    return sanitize_name(*args, **kwargs)


  # Helpers to find media

  # Match heard string to something in the results
  def matchHeard(self, heard, results, lookingFor='label'):
    located = None

    heard_lower = heard.lower()

    # Very ugly hack for German Alexa.  In English, if a user specifies
    # 'percent', she converts it to a '%' symbol.  In German, for whatever
    # reason, she leaves it unconverted as 'prozent'.  Let's convert here.
    heard_lower = re.sub(r'prozent(?=[.,\s]|$)', '%', heard_lower)

    print 'Trying to match: ' + heard_lower.encode("utf-8")

    heard_ascii = sanitize_name(heard_lower)
    for result in results:
      result_lower = result[lookingFor].lower()

      # Direct comparison
      if type(heard_lower) is type(result_lower):
        if result_lower == heard_lower:
          print 'Simple match on direct comparison'
          located = result
          break

      # Strip out non-ascii symbols
      result_name = sanitize_name(result_lower)

      # Direct comparison (ASCII)
      if result_name == heard_ascii:
        print 'Simple match on direct comparison (ASCII)'
        located = result
        break

    if not located:
      print 'Simple match failed, trying fuzzy match...'
      print 'Processing %d items...' % (len(results))

      fuzzy_results = []
      for f in (None, digits2roman, words2roman, words2digits, digits2words):
        try:
          if f is not None:
            ms = f(heard_lower, self.language)
            # no need to bother fuzzy matching if the string didn't change
            if ms == heard_lower: continue
            mf = f.__name__
          else:
            ms = heard_lower
            mf = 'heard'

          print '  %s: "%s"' % (mf, ms.encode("utf-8"))

          rv = process.extractOne(ms, [d[lookingFor] for d in results], scorer=fuzz.UQRatio, score_cutoff=75)
          if rv:
            fuzzy_results.append(rv)
            print '   -- Score %d%%' % (rv[1])
            if rv[1] == 95:
              # Let's consider a 95% match 'good enough'
              break
        except:
          continue

      # Got a match?
      if len(fuzzy_results) > 0:
        winner = sorted(fuzzy_results, key=lambda x: x[1], reverse=True)[0]
        print '  WINNER: "%s" @ %d%%' % (winner[0].encode("utf-8"), winner[1])
        located = (item for item in results if item[lookingFor] == winner[0]).next()

    return located


  def FindVideoPlaylist(self, heard_search):
    print 'Searching for video playlist "%s"' % (heard_search.encode("utf-8"))

    playlists = self.GetVideoPlaylists()
    if 'result' in playlists and 'files' in playlists['result']:
      playlists_list = playlists['result']['files']
      located = self.matchHeard(heard_search, playlists_list, 'label')

      if located:
        print 'Located video playlist "%s"' % (located['label'].encode("utf-8"))
        return located['file'], located['label']

    return None, None


  def FindAudioPlaylist(self, heard_search):
    print 'Searching for audio playlist "%s"' % (heard_search.encode("utf-8"))

    playlists = self.GetMusicPlaylists()
    if 'result' in playlists and 'files' in playlists['result']:
      playlists_list = playlists['result']['files']
      located = self.matchHeard(heard_search, playlists_list, 'label')

      if located:
        print 'Located audio playlist "%s"' % (located['label'].encode("utf-8"))
        return located['file'], located['label']

    return None, None


  def FindVideoGenre(self, heard_search, genretype='movie'):
    print 'Searching for movie genre "%s"' % (heard_search.encode("utf-8"))

    genres = self.GetVideoGenres(genretype)
    if 'result' in genres and 'genres' in genres['result']:
      genres_list = genres['result']['genres']
      located = self.matchHeard(heard_search, genres_list, 'label')

      if located:
        print 'Located genre "%s"' % (located['label'].encode("utf-8"))
        return located['genreid'], located['label']

    return None, None


  def FindMovie(self, heard_search):
    print 'Searching for movie "%s"' % (heard_search.encode("utf-8"))

    movies = self.GetMovies()
    if 'result' in movies and 'movies' in movies['result']:
      movies_array = movies['result']['movies']
      located = self.matchHeard(heard_search, movies_array)

      if located:
        print 'Located movie "%s"' % (located['label'].encode("utf-8"))
        return located['movieid'], located['label']

    return None, None


  def FindTvShow(self, heard_search):
    print 'Searching for show "%s"' % (heard_search.encode("utf-8"))

    shows = self.GetTvShows()
    if 'result' in shows and 'tvshows' in shows['result']:
      shows_array = shows['result']['tvshows']
      located = self.matchHeard(heard_search, shows_array)

      if located:
        print 'Located tvshow "%s"' % (located['label'].encode("utf-8"))
        return located['tvshowid'], located['label']

    return None, None


  def FindArtist(self, heard_search):
    print 'Searching for artist "%s"' % (heard_search.encode("utf-8"))

    artists = self.GetMusicArtists()
    if 'result' in artists and 'artists' in artists['result']:
      artists_list = artists['result']['artists']
      located = self.matchHeard(heard_search, artists_list, 'artist')

      if located:
        print 'Located artist "%s"' % (located['label'].encode("utf-8"))
        return located['artistid'], located['label']

    return None, None


  def FindAlbum(self, heard_search, artist_id=None):
    print 'Searching for album "%s"' % (heard_search.encode("utf-8"))

    if artist_id:
      albums = self.GetArtistAlbums(artist_id)
    else:
      albums = self.GetAlbums()
    if 'result' in albums and 'albums' in albums['result']:
      albums_list = albums['result']['albums']
      located = self.matchHeard(heard_search, albums_list, 'label')

      if located:
        print 'Located album "%s"' % (located['label'].encode("utf-8"))
        return located['albumid'], located['label']

    return None, None


  def FindSong(self, heard_search, artist_id=None):
    print 'Searching for song "%s"' % (heard_search.encode("utf-8"))

    if artist_id:
      songs = self.GetArtistSongs(artist_id)
    else:
      songs = self.GetSongs()
    if 'result' in songs and 'songs' in songs['result']:
      songs_list = songs['result']['songs']
      located = self.matchHeard(heard_search, songs_list, 'label')

      if located:
        print 'Located song "%s"' % (located['label'].encode("utf-8"))
        return located['songid'], located['label']

    return None, None


  def FindAddon(self, heard_search):
    print 'Searching for addon "%s"' % (heard_search.encode("utf-8"))

    for content in ['video', 'audio', 'image', 'executable']:
      addons = self.GetAddons(content)
      if 'result' in addons and 'addons' in addons['result']:
        addons_list = addons['result']['addons']
        located = self.matchHeard(heard_search, addons_list, 'name')

        if located:
          print 'Located addon "%s"' % (located['name'].encode("utf-8"))
          return located['addonid'], located['name']

    return None, None


  # Playlists

  def ClearAudioPlaylist(self):
    return self.SendCommand(RPCString("Playlist.Clear", {"playlistid": 0}))


  def AddSongToPlaylist(self, song_id):
    return self.SendCommand(RPCString("Playlist.Add", {"playlistid": 0, "item": {"songid": int(song_id)}}))


  def AddSongsToPlaylist(self, song_ids, shuffle=False):
    songs_array = []

    for song_id in song_ids:
      temp_song = {}
      temp_song['songid'] = song_id
      songs_array.append(temp_song)

    if shuffle:
      random.shuffle(songs_array)

    # Segment the requests into chunks that Kodi will accept in a single call
    song_groups = [songs_array[x:x+2000] for x in range(0, len(songs_array), 2000)]
    for a in song_groups:
      print "Adding %d items to the queue..." % (len(a))
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


  # Note that subsequent shuffle commands won't work with this, as Kodi
  # considers a playlist to be a single item.
  def StartAudioPlaylist(self, playlist_file=None):
    if playlist_file is not None and playlist_file != '':
      return self.SendCommand(RPCString("Player.Open", {"item": {"file": playlist_file}}))
    else:
      return self.SendCommand(RPCString("Player.Open", {"item": {"playlistid": 0}}))


  def ClearVideoPlaylist(self):
    return self.SendCommand(RPCString("Playlist.Clear", {"playlistid": 1}))


  def AddEpisodeToPlayList(self, ep_id):
    return self.SendCommand(RPCString("Playlist.Add", {"playlistid": 1, "item": {"episodeid": int(ep_id)}}))


  def AddEpisodesToPlaylist(self, episode_ids, shuffle=False):
    episodes_array = []

    for episode_id in episode_ids:
      temp_episode = {}
      temp_episode['episodeid'] = episode_id
      episodes_array.append(temp_episode)

    if shuffle:
      random.shuffle(episodes_array)

    # Segment the requests into chunks that Kodi will accept in a single call
    episode_groups = [episodes_array[x:x+2000] for x in range(0, len(episodes_array), 2000)]
    for a in episode_groups:
      print "Adding %d items to the queue..." % (len(a))
      res = self.SendCommand(RPCString("Playlist.Add", {"playlistid": 1, "item": a}))

    return res


  def AddMovieToPlaylist(self, movie_id):
    return self.SendCommand(RPCString("Playlist.Add", {"playlistid": 1, "item": {"movieid": int(movie_id)}}))


  def AddVideosToPlaylist(self, video_files, shuffle=False):
    videos_array = []

    for video_file in video_files:
      temp_video = {}
      temp_video['file'] = video_file
      videos_array.append(temp_video)

    if shuffle:
      random.shuffle(videos_array)

    # Segment the requests into chunks that Kodi will accept in a single call
    video_groups = [videos_array[x:x+2000] for x in range(0, len(videos_array), 2000)]
    for a in video_groups:
      print "Adding %d items to the queue..." % (len(a))
      res = self.SendCommand(RPCString("Playlist.Add", {"playlistid": 1, "item": a}))

    return res


  def GetVideoPlaylistItems(self):
    return self.SendCommand(RPCString("Playlist.GetItems", {"playlistid": 1}))


  # Note that subsequent shuffle commands won't work with this, as Kodi
  # considers a playlist to be a single item.
  def StartVideoPlaylist(self, playlist_file=None):
    if playlist_file is not None and playlist_file != '':
      return self.SendCommand(RPCString("Player.Open", {"item": {"file": playlist_file}}))
    else:
      return self.SendCommand(RPCString("Player.Open", {"item": {"playlistid": 1}}))


  # Direct plays

  def PlayEpisode(self, ep_id, resume=True):
    return self.SendCommand(RPCString("Player.Open", {"item": {"episodeid": ep_id}, "options": {"resume": resume}}))


  def PlayMovie(self, movie_id, resume=True):
    return self.SendCommand(RPCString("Player.Open", {"item": {"movieid": movie_id}, "options": {"resume": resume}}))


  def PartyPlayMusic(self):
    return self.SendCommand(RPCString("Player.Open", {"item": {"partymode": "music"}}))


  # Tell Kodi to update its video or music libraries

  def UpdateVideo(self):
    return self.SendCommand(RPCString("VideoLibrary.Scan"))


  def CleanVideo(self):
    return self.SendCommand(RPCString("VideoLibrary.Clean"))


  def UpdateMusic(self):
    return self.SendCommand(RPCString("AudioLibrary.Scan"))


  def CleanMusic(self):
    return self.SendCommand(RPCString("AudioLibrary.Clean"))


  # Perform UI actions that match the normal remote control buttons

  def PageUp(self):
    return self.SendCommand(RPCString("Input.ExecuteAction", {"action":"pageup"}))


  def PageDown(self):
    return self.SendCommand(RPCString("Input.ExecuteAction", {"action":"pagedown"}))


  def ToggleWatched(self):
    return self.SendCommand(RPCString("Input.ExecuteAction", {"action":"togglewatched"}))


  def Info(self):
    return self.SendCommand(RPCString("Input.Info"))


  def Menu(self):
    return self.SendCommand(RPCString("Input.ContextMenu"))


  def Home(self):
    return self.SendCommand(RPCString("Input.Home"))


  def Select(self):
    return self.SendCommand(RPCString("Input.Select"))


  def Up(self):
    return self.SendCommand(RPCString("Input.Up"))


  def Down(self):
    return self.SendCommand(RPCString("Input.Down"))


  def Left(self):
    return self.SendCommand(RPCString("Input.Left"))


  def Right(self):
    return self.SendCommand(RPCString("Input.Right"))


  def Back(self):
    return self.SendCommand(RPCString("Input.Back"))


  def ToggleFullscreen(self):
    return self.SendCommand(RPCString("GUI.SetFullscreen", {"fullscreen":"toggle"}))


  def ToggleMute(self):
    return self.SendCommand(RPCString("Application.SetMute", {"mute":"toggle"}))


  def GetCurrentVolume(self):
    return self.SendCommand(RPCString("Application.GetProperties", {"properties":["volume", "muted"]}))


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
    return self.SendCommand(RPCString("Application.SetVolume", {"volume":vol}))


  def VolumeDown(self):
    resp = self.GetCurrentVolume()
    vol = resp['result']['volume']
    if vol % 10 != 0:
      # round up to nearest 10 first
      vol -= vol % -10
    vol -= 10
    if vol < 0:
      vol = 0
    return self.SendCommand(RPCString("Application.SetVolume", {"volume":vol}))


  def VolumeSet(self, vol, percent=True):
    if vol < 0:
      vol = 0
    if not percent:
      # specified with scale of 0 to 10
      vol *= 10
    if vol > 100:
      vol = 100
    return self.SendCommand(RPCString("Application.SetVolume", {"volume":vol}))


  # Player controls

  def PlayerPlayPause(self):
    playerid = self.GetPlayerID()
    if playerid is not None:
      return self.SendCommand(RPCString("Player.PlayPause", {"playerid":playerid}))


  def PlayerSkip(self):
    playerid = self.GetPlayerID()
    if playerid is not None:
      return self.SendCommand(RPCString("Player.GoTo", {"playerid":playerid, "to": "next"}))


  def PlayerPrev(self):
    playerid = self.GetPlayerID()
    if playerid is not None:
      self.SendCommand(RPCString("Player.GoTo", {"playerid":playerid, "to": "previous"}))
      return self.SendCommand(RPCString("Player.GoTo", {"playerid":playerid, "to": "previous"}))


  def PlayerStartOver(self):
    playerid = self.GetPlayerID()
    if playerid is not None:
      return self.SendCommand(RPCString("Player.Seek", {"playerid":playerid, "value": 0}))


  def PlayerStop(self):
    playerid = self.GetPlayerID()
    if playerid is not None:
      return self.SendCommand(RPCString("Player.Stop", {"playerid":playerid}))


  def PlayerSeek(self, seconds):
    playerid = self.GetPlayerID()
    if playerid:
      return self.SendCommand(RPCString("Player.Seek", {"playerid":playerid, "value":{"seconds":seconds}}))


  def PlayerSeekSmallForward(self):
    playerid = self.GetPlayerID()
    if playerid:
      return self.SendCommand(RPCString("Player.Seek", {"playerid":playerid, "value":"smallforward"}))


  def PlayerSeekSmallBackward(self):
    playerid = self.GetPlayerID()
    if playerid:
      return self.SendCommand(RPCString("Player.Seek", {"playerid":playerid, "value":"smallbackward"}))


  def PlayerSeekBigForward(self):
    playerid = self.GetPlayerID()
    if playerid:
      return self.SendCommand(RPCString("Player.Seek", {"playerid":playerid, "value":"bigforward"}))


  def PlayerSeekBigBackward(self):
    playerid = self.GetPlayerID()
    if playerid:
      return self.SendCommand(RPCString("Player.Seek", {"playerid":playerid, "value":"bigbackward"}))


  def PlayerShuffleOn(self):
    playerid = self.GetPlayerID()
    if playerid is not None:
      return self.SendCommand(RPCString("Player.SetShuffle", {"playerid":playerid, "shuffle":True}))


  def PlayerShuffleOff(self):
    playerid = self.GetPlayerID()
    if playerid is not None:
      return self.SendCommand(RPCString("Player.SetShuffle", {"playerid":playerid, "shuffle":False}))


  def PlayerLoopOn(self):
    playerid = self.GetPlayerID()
    if playerid is not None:
      return self.SendCommand(RPCString("Player.SetRepeat", {"playerid":playerid, "repeat":"cycle"}))


  def PlayerLoopOff(self):
    playerid = self.GetPlayerID()
    if playerid is not None:
      return self.SendCommand(RPCString("Player.SetRepeat", {"playerid":playerid, "repeat":"off"}))


  def PlayerSubtitlesOn(self):
    playerid = self.GetVideoPlayerID()
    if playerid:
      return self.SendCommand(RPCString("Player.SetSubtitle", {"playerid":playerid, "subtitle":"on"}))


  def PlayerSubtitlesOff(self):
    playerid = self.GetVideoPlayerID()
    if playerid:
      return self.SendCommand(RPCString("Player.SetSubtitle", {"playerid":playerid, "subtitle":"off"}))


  def PlayerSubtitlesNext(self):
    playerid = self.GetVideoPlayerID()
    if playerid:
      return self.SendCommand(RPCString("Player.SetSubtitle", {"playerid":playerid, "subtitle":"next", "enable":True}))


  def PlayerSubtitlesPrevious(self):
    playerid = self.GetVideoPlayerID()
    if playerid:
      return self.SendCommand(RPCString("Player.SetSubtitle", {"playerid":playerid, "subtitle":"previous", "enable":True}))


  def PlayerAudioStreamNext(self):
    playerid = self.GetVideoPlayerID()
    if playerid:
      return self.SendCommand(RPCString("Player.SetAudioStream", {"playerid":playerid, "stream":"next"}))


  def PlayerAudioStreamPrevious(self):
    playerid = self.GetVideoPlayerID()
    if playerid:
      return self.SendCommand(RPCString("Player.SetAudioStream", {"playerid":playerid, "stream":"previous"}))


  def PlayerMoveUp(self):
    playerid = self.GetPicturePlayerID()
    if playerid:
      return self.SendCommand(RPCString("Player.Move", {"playerid":playerid, "direction":"up"}))


  def PlayerMoveDown(self):
    playerid = self.GetPicturePlayerID()
    if playerid:
      return self.SendCommand(RPCString("Player.Move", {"playerid":playerid, "direction":"down"}))


  def PlayerMoveLeft(self):
    playerid = self.GetPicturePlayerID()
    if playerid:
      return self.SendCommand(RPCString("Player.Move", {"playerid":playerid, "direction":"left"}))


  def PlayerMoveRight(self):
    playerid = self.GetPicturePlayerID()
    if playerid:
      return self.SendCommand(RPCString("Player.Move", {"playerid":playerid, "direction":"right"}))


  def PlayerZoom(self, lvl=0):
    playerid = self.GetPicturePlayerID()
    if playerid and lvl > 0 and lvl < 11:
      return self.SendCommand(RPCString("Player.Zoom", {"playerid":playerid, "zoom":lvl}))


  def PlayerZoomIn(self):
    playerid = self.GetPicturePlayerID()
    if playerid:
      return self.SendCommand(RPCString("Player.Zoom", {"playerid":playerid, "zoom":"in"}))


  def PlayerZoomOut(self):
    playerid = self.GetPicturePlayerID()
    if playerid:
      return self.SendCommand(RPCString("Player.Zoom", {"playerid":playerid, "zoom":"out"}))


  def PlayerRotateClockwise(self):
    playerid = self.GetPicturePlayerID()
    if playerid:
      return self.SendCommand(RPCString("Player.Rotate", {"playerid":playerid, "value":"clockwise"}))


  def PlayerRotateCounterClockwise(self):
    playerid = self.GetPicturePlayerID()
    if playerid:
      return self.SendCommand(RPCString("Player.Rotate", {"playerid":playerid, "value":"counterclockwise"}))


  # Addons

  def AddonExecute(self, addon_id, params={}):
    return self.SendCommand(RPCString("Addons.ExecuteAddon", {"addonid":addon_id, "params":params}))


  def AddonGlobalSearch(self, needle=''):
    return self.AddonExecute("script.globalsearch", {"searchstring":needle.encode("utf-8")})


  def AddonCinemaVision(self):
    return self.AddonExecute("script.cinemavision", ["experience"])


  # Library queries

  # content can be: video, audio, image, executable, or unknown
  def GetAddons(self, content):
    if content:
      return self.SendCommand(RPCString("Addons.GetAddons", {"content":content, "properties":["name"]}))
    else:
      return self.SendCommand(RPCString("Addons.GetAddons", {"properties":["name"]}))


  def GetAddonDetails(self, addon_id):
    return self.SendCommand(RPCString("Addons.GetAddonDetails", {"addonid":addon_id, "properties":["name", "version", "description", "summary"]}))


  # mediatype should be one of:
  # movies, tvshows, episodes, musicvideos, artists, albums, songs
  #
  # if mediatype == 'media', it will recommend one item from the whole library
  def GetRecommendedItem(self, mediatype='media'):
    if mediatype == 'movies' or mediatype == 'episodes':
      action = 'inprogressandrecommended'
    else:
      action = 'recommended'

    shw_path = 'plugin://script.skin.helper.widgets/?action=%s&mediatype=%s' % (action, mediatype)
    data = self.SendCommand(RPCString("Files.GetDirectory", {"directory": shw_path}))

    answer = []
    if 'files' in data['result']:
      # shuffle so if we get called more than once before skin helper updates
      # the list, we have a chance at a different item
      random.shuffle(data['result']['files'])
      m = data['result']['files'][0]

      answer.append(m['type'])
      answer.append(m['label'])
      if m['type'] == 'artist':
        # script doesn't provide artist id as a separate list item
        m_id = m['file'].rsplit('/', 1)[1].strip()
      elif m['type'] == 'album':
        # script doesn't provide album id as a separate list item
        m_id = m['file'].rsplit('=', 1)[1].strip()
      else:
        m_id = m['id']
      answer.append(int(m_id))
    elif mediatype == 'songs':
      # if skin helper doesn't return anything for recommended songs, just
      # return an empty list and caller can do something else
      return ['song', '', 0]

    return answer


  def GetRecommendedVideoItem(self):
    answer = []
    items = []
    for content in ['movies', 'tvshows', 'episodes', 'musicvideos']:
      item = self.GetRecommendedItem(content)
      if len(item) > 0:
        items.append(item)

    if len(items) > 0:
      answer = random.choice(items)

    return answer


  def GetRecommendedAudioItem(self):
    answer = []
    items = []
    for content in ['musicvideos', 'artists', 'albums', 'songs']:
      item = self.GetRecommendedItem(content)
      if len(item) > 0:
        items.append(item)

    if len(items) > 0:
      answer = random.choice(items)

    return answer


  def GetPlaylistItems(self, playlist_file):
    return self.SendCommand(RPCString("Files.GetDirectory", {"directory": playlist_file}))


  def GetMusicPlaylists(self):
    return self.SendCommand(RPCString("Files.GetDirectory", {"directory": "special://musicplaylists"}))


  def GetMusicArtists(self):
    return self.SendCommand(RPCString("AudioLibrary.GetArtists", {"albumartistsonly": False}))


  def GetMusicGenres(self):
    return self.SendCommand(RPCString("AudioLibrary.GetGenres"))


  def GetArtistAlbums(self, artist_id):
    return self.SendCommand(RPCString("AudioLibrary.GetAlbums", {"filter": {"artistid": int(artist_id)}}))


  def GetNewestAlbumFromArtist(self, artist_id):
    data = self.SendCommand(RPCString("AudioLibrary.GetAlbums", {"limits": {"end": 1}, "filter": {"artistid": int(artist_id)}, "sort": {"method": "dateadded", "order": "descending"}}))
    if 'albums' in data['result']:
      album = data['result']['albums'][0]
      return album['albumid']
    else:
      return None


  def GetAlbums(self):
    return self.SendCommand(RPCString("AudioLibrary.GetAlbums"))


  def GetArtistSongs(self, artist_id):
    return self.SendCommand(RPCString("AudioLibrary.GetSongs", {"filter": {"artistid": int(artist_id)}}))


  def GetArtistSongsPath(self, artist_id):
    return self.SendCommand(RPCString("AudioLibrary.GetSongs", {"filter": {"artistid": int(artist_id)}, "properties":["file"]}))


  def GetAlbumDetails(self, album_id):
    data = self.SendCommand(RPCString("AudioLibrary.GetAlbumDetails", {"albumid": int(album_id), "properties":["artist"]}))
    return data['result']['albumdetails']


  def GetAlbumSongs(self, album_id):
    return self.SendCommand(RPCString("AudioLibrary.GetSongs", {"filter": {"albumid": int(album_id)}}))


  def GetAlbumSongsPath(self, album_id):
    return self.SendCommand(RPCString("AudioLibrary.GetSongs", {"filter": {"albumid": int(album_id)}, "properties":["file"]}))


  def GetSongs(self):
    return self.SendCommand(RPCString("AudioLibrary.GetSongs"))


  def GetSongsPath(self):
    return self.SendCommand(RPCString("AudioLibrary.GetSongs", {"properties":["file"]}))


  def GetSongIdPath(self, song_id):
    return self.SendCommand(RPCString("AudioLibrary.GetSongDetails", {"songid": int(song_id), "properties":["file"]}))


  def GetSongDetails(self, song_id):
    data = self.SendCommand(RPCString("AudioLibrary.GetSongDetails", {"songid": int(song_id), "properties":["artist"]}))
    return data['result']['songdetails']


  def GetRecentlyAddedAlbums(self):
    return self.SendCommand(RPCString("AudioLibrary.GetRecentlyAddedAlbums", {'properties':['artist']}))


  def GetRecentlyAddedSongs(self):
    return self.SendCommand(RPCString("AudioLibrary.GetRecentlyAddedSongs", {'properties':['artist']}))


  def GetRecentlyAddedSongsPath(self):
    return self.SendCommand(RPCString("AudioLibrary.GetRecentlyAddedSongs", {'properties':['artist', 'file']}))


  def GetVideoPlaylists(self):
    return self.SendCommand(RPCString("Files.GetDirectory", {"directory": "special://videoplaylists"}))


  def GetTvShowDetails(self, show_id):
    data = self.SendCommand(RPCString("VideoLibrary.GetTVShowDetails", {'tvshowid':show_id, 'properties':['art']}))
    return data['result']['tvshowdetails']


  def GetTvShows(self):
    return self.SendCommand(RPCString("VideoLibrary.GetTVShows"))


  def GetMovieDetails(self, movie_id):
    data = self.SendCommand(RPCString("VideoLibrary.GetMovieDetails", {'movieid':movie_id, 'properties':['resume']}))
    return data['result']['moviedetails']


  def GetMovies(self):
    return self.SendCommand(RPCString("VideoLibrary.GetMovies"))


  def GetMoviesByGenre(self, genre):
    return self.SendCommand(RPCString("VideoLibrary.GetMovies", {"filter":{"genre":genre}}))


  def GetVideoGenres(self, genretype='movie'):
    return self.SendCommand(RPCString("VideoLibrary.GetGenres", {"type":genretype}))


  def GetEpisodeDetails(self, ep_id):
    data = self.SendCommand(RPCString("VideoLibrary.GetEpisodeDetails", {"episodeid": int(ep_id), "properties":["showtitle", "season", "episode", "resume"]}))
    return data['result']['episodedetails']


  def GetEpisodesFromShow(self, show_id):
    return self.SendCommand(RPCString("VideoLibrary.GetEpisodes", {"tvshowid": int(show_id)}))


  def GetUnwatchedEpisodesFromShow(self, show_id):
    return self.SendCommand(RPCString("VideoLibrary.GetEpisodes", {"tvshowid": int(show_id), "filter":{"field":"playcount", "operator":"lessthan", "value":"1"}}))


  def GetNewestEpisodeFromShow(self, show_id):
    data = self.SendCommand(RPCString("VideoLibrary.GetEpisodes", {"limits":{"end":1},"tvshowid": int(show_id), "sort":{"method":"dateadded", "order":"descending"}}))
    if 'episodes' in data['result']:
      episode = data['result']['episodes'][0]
      return episode['episodeid']
    else:
      return None


  def GetNextUnwatchedEpisode(self, show_id):
    data = self.SendCommand(RPCString("VideoLibrary.GetEpisodes", {"limits":{"end":1},"tvshowid": int(show_id), "filter":{"field":"playcount", "operator":"lessthan", "value":"1"}, "properties":["playcount"]}))
    if 'episodes' in data['result']:
      episode = data['result']['episodes'][0]
      return episode['episodeid']
    else:
      return None


  def GetLastWatchedShow(self):
    return self.SendCommand(RPCString("VideoLibrary.GetEpisodes", {"limits":{"end":1}, "filter":{"field":"playcount", "operator":"greaterthan", "value":"0"}, "filter":{"field":"lastplayed", "operator":"greaterthan", "value":"0"}, "sort":{"method":"lastplayed", "order":"descending"}, "properties":["tvshowid", "showtitle"]}))


  def GetSpecificEpisode(self, show_id, season, episode):
    data = self.SendCommand(RPCString("VideoLibrary.GetEpisodes", {"tvshowid": int(show_id), "season": int(season), "properties": ["season", "episode"]}))
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
    return self.SendCommand(RPCString("VideoLibrary.GetEpisodes", {"tvshowid": int(show_id), "properties": ["season", "episode"]}))


  # Returns a list of dictionaries with information about episodes that have been watched.
  def GetWatchedEpisodes(self):
    return self.SendCommand(RPCString("VideoLibrary.GetEpisodes", {"filter":{"field":"playcount", "operator":"greaterthan", "value":"0"}, "properties":["playcount", "showtitle", "season", "episode", "lastplayed" ]}))


  # Returns a list of dictionaries with information about unwatched movies. Useful for
  # telling/showing users what's ready to be watched. Setting max to very high values
  # can take a long time.
  def GetUnwatchedMovies(self):
    data = self.SendCommand(RPCString("VideoLibrary.GetMovies", {"limits":{"end":self.max_unwatched_movies}, "filter":{"field":"playcount", "operator":"lessthan", "value":"1"}, "sort":{"method":"dateadded", "order":"descending"}, "properties":["title", "playcount", "dateadded"]}))
    answer = []
    if 'movies' in data['result']:
      for d in data['result']['movies']:
        answer.append({'title':d['title'], 'movieid':d['movieid'], 'label':d['label'], 'dateadded':datetime.datetime.strptime(d['dateadded'], "%Y-%m-%d %H:%M:%S")})
    return answer

  # Returns a list of dictionaries with information about unwatched movies in a particular genre. Useful for
  # telling/showing users what's ready to be watched. Setting max to very high values
  # can take a long time.
  def GetUnwatchedMoviesByGenre(self, genre):
    data = self.SendCommand(RPCString("VideoLibrary.GetMovies", {"limits":{"end":self.max_unwatched_movies}, "filter":{"and":[{"field":"playcount", "operator":"lessthan", "value":"1"}, {"field":"genre", "operator":"contains", "value":genre}]}, "sort":{"method":"dateadded", "order":"descending"}, "properties":["title", "playcount", "dateadded"]}))
    answer = []
    if 'movies' in data['result']:
      for d in data['result']['movies']:
        answer.append({'title':d['title'], 'movieid':d['movieid'], 'label':d['label'], 'dateadded':datetime.datetime.strptime(d['dateadded'], "%Y-%m-%d %H:%M:%S")})
    return answer


  # Returns a list of dictionaries with information about unwatched episodes. Useful for
  # telling/showing users what's ready to be watched. Setting max to very high values
  # can take a long time.
  def GetUnwatchedEpisodes(self):
    data = self.SendCommand(RPCString("VideoLibrary.GetEpisodes", {"limits":{"end":self.max_unwatched_episodes}, "filter":{"field":"playcount", "operator":"lessthan", "value":"1"}, "sort":{"method":"dateadded", "order":"descending"}, "properties":["title", "playcount", "showtitle", "tvshowid", "dateadded" ]}))
    answer = []
    if 'episodes' in data['result']:
      shows = set([d['tvshowid'] for d in data['result']['episodes']])
      show_info = {}
      for show in shows:
        show_info[show] = self.GetTvShowDetails(show_id=show)
      for d in data['result']['episodes']:
        showinfo = show_info[d['tvshowid']]
        answer.append({'title':d['title'], 'episodeid':d['episodeid'], 'show':d['showtitle'], 'label':d['label'], 'dateadded':datetime.datetime.strptime(d['dateadded'], "%Y-%m-%d %H:%M:%S")})
    return answer


  def GetMusicVideoDetails(self, mv_id):
    data = self.SendCommand(RPCString("VideoLibrary.GetMusicVideoDetails", {"musicvideoid": int(mv_id), "properties": ["artist"]}))
    return data['result']['musicvideodetails']


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

    if self.config.get(self.deviceId, 'user_proxy') in accepted_answers:
      stream_url = 'https://kodi-music-proxy.herokuapp.com/proxy?file=' + url
    elif self.config.get(self.deviceId, 'alt_proxy'):
      stream_url = os.getenv('ALT_PROXY') + self.config.get(self.deviceId, 'alt_proxy')
    else:
      stream_url = url

    return stream_url

  # Get the first active player.
  def GetPlayerID(self, playertype=['picture', 'audio', 'video']):
    data = self.SendCommand(RPCString("Player.GetActivePlayers"))
    result = data.get("result", [])
    if len(result) > 0:
      for curitem in result:
        if curitem.get("type") in playertype:
          return curitem.get("playerid")
    return None

  # Get the first active Video player.
  def GetVideoPlayerID(self, playertype=['video']):
    data = self.SendCommand(RPCString("Player.GetActivePlayers"))
    result = data.get("result", [])
    if len(result) > 0:
      for curitem in result:
        if curitem.get("type") in playertype:
          return curitem.get("playerid")
    return None


  # Get the first active Audio player.
  def GetAudioPlayerID(self, playertype=['audio']):
    data = self.SendCommand(RPCString("Player.GetActivePlayers"))
    result = data.get("result", [])
    if len(result) > 0:
      for curitem in result:
        if curitem.get("type") in playertype:
          return curitem.get("playerid")
    return None


  # Get the first active Picture player.
  def GetPicturePlayerID(self, playertype=['picture']):
    data = self.SendCommand(RPCString("Player.GetActivePlayers"))
    result = data.get("result", [])
    if len(result) > 0:
      for curitem in result:
        if curitem.get("type") in playertype:
          return curitem.get("playerid")
    return None


  # Information about the video or audio that's currently playing

  def GetActivePlayItem(self):
    playerid = self.GetPlayerID()
    if playerid is not None:
      data = self.SendCommand(RPCString("Player.GetItem", {"playerid":playerid, "properties":["title", "album", "artist", "season", "episode", "showtitle", "tvshowid", "description"]}))
      #print data['result']['item']
      return data['result']['item']


  def GetActivePlayProperties(self):
    playerid = self.GetPlayerID()
    if playerid is not None:
      data = self.SendCommand(RPCString("Player.GetProperties", {"playerid":playerid, "properties":["currentaudiostream", "currentsubtitle", "canshuffle", "shuffled", "canrepeat", "repeat", "canzoom", "canrotate", "canmove"]}))
      #print data['result']
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
      data = self.SendCommand(RPCString("Player.GetProperties", {"playerid":playerid, "properties":["percentage", "speed", "time", "totaltime"]}))
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
        return {'state':'play' if speed > 0 else 'pause', 'time':cur, 'time_hours':hours_cur, 'time_mins':mins_cur, 'totaltime':total, 'total_hours':hours_total, 'total_mins':mins_total, 'pct':data['result']['percentage']}
    return {'state':'stop'}
