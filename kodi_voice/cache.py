#!/usr/bin/env python

import json
import requests
import boto3
import botocore
import owncloud
import hashlib
import io
import logging

log = logging.getLogger(__name__)


class KodiCache():
  def __init__(self, bucket_name=None, **kwargs):
    self.enabled = False
    self.backend = None
    self.s3 = None
    self.oc = None

    log.info('Initalizing')

    # Amazon S3 bucket or directory name
    self.bucket_name = bucket_name

    # Amazon credentials for accessing an S3 bucket
    self.aws_secret_access_key = kwargs.get('aws_secret_access_key', None)
    self.aws_access_key_id = kwargs.get('aws_access_key_id', None)

    # ownCloud/nextCloud credentials
    self.oc_url = kwargs.get('oc_url', None)
    self.oc_user = kwargs.get('oc_user', None)
    self.oc_pass = kwargs.get('oc_password', None)

    if self.bucket_name:
      # Amazon S3
      if self.aws_secret_access_key and self.aws_access_key_id:
        self.s3 = boto3.resource('s3', aws_secret_access_key=self.aws_secret_access_key, aws_access_key_id=self.aws_access_key_id)

        log.info('Accessing bucket %s', self.bucket_name)
        self.bucket = self.s3.Bucket(self.bucket_name)

        try:
          self.s3.meta.client.head_bucket(Bucket=self.bucket_name)
        except botocore.exceptions.ClientError as e:
          log.error('Error %s accessing bucket %s', e.response['Error']['Code'], self.bucket_name)
          # continue on without the cache
          pass
        else:
          self.enabled = True
          self.backend = 'Amazon S3'

      # ownCloud/nextCloud
      elif self.oc_url and self.oc_user and self.oc_pass:
        self.oc = owncloud.Client(self.oc_url)
        self.oc.login(self.oc_user, self.oc_pass)

        if not self.bucket_name[0] == '/':
          self.bucket_name = '/' + self.bucket_name

        file_info = None
        try:
          file_info = self.oc.file_info(self.bucket_name)
        except owncloud.HTTPResponseError as e:
          if e.status_code == 404:
            if self.oc.mkdir(self.bucket_name):
              file_info = self.oc.file_info(self.bucket_name)
              pass
            else:
              log.error('Could not create cache directory %s', self.bucket_name)
          else:
            log.error('Error %d accessing directory %s', e.status_code, self.bucket_name)
            # continue on without the cache
            pass

        if isinstance(file_info, owncloud.FileInfo) and not file_info.is_dir():
          log.error('%s exists, but is not a directory!', self.bucket_name)
        else:
          self.enabled = True
          self.backend = 'ownCloud'

    if self.enabled:
      assert self.backend is not None

      if log.getEffectiveLevel() == logging.DEBUG:
        cache_objs = self.ls()
        if not cache_objs:
          log.debug('Object cache empty')
        else:
          log.debug('Objects in cache:')
          for o in cache_objs:
            log.debug('  %s', o)

      log.info('Initialized using %s cache backend', self.backend)
    else:
      log.info('Disabled')

  def ls(self):
    listing = []
    if self.enabled:
      if self.s3:
        listing = [key.key for key in self.bucket.objects.all()]
      elif self.oc:
        listing = [f.get_name() for f in self.oc.list(self.bucket_name)]
    return listing

  def clear(self):
    if self.enabled:
      log.debug('Clearing cache objects')

      if self.s3:
        for key in self.bucket.objects.all():
          key.delete()
        #self.bucket.delete()
      elif self.oc:
        for f in self.oc.list(self.bucket_name):
          self.oc.delete(self.bucket_name + '/' + f.get_name())
        self.oc.delete(self.bucket_name)

      log.info('Cleared all cache objects')

  def add(self, cache_file, url, auth, command, timeout, wait_resp=True):
    try:
      # Fetch the response from Kodi
      r = requests.post(url, data=command, auth=auth, timeout=timeout)
    except requests.exceptions.ReadTimeout:
      if not wait_resp:
        # Caller doesn't care about the response anyway -- this is mostly for
        # Player.Open and other methods that can either never fail or we don't
        # respond any differently if they do.
        pass
      else:
        raise
    else:
      if r.encoding is None:
        r.encoding = 'utf-8'

      try:
        resp = r.json()
      except:
        log.error('JSON decoding failed {}'.format(r))
        raise

      if self.enabled and cache_file:
        log.debug('Adding object %s', cache_file)

        try:
          if self.s3:
            self.s3.Object(self.bucket_name, cache_file).put(Body=json.dumps(resp))
          elif self.oc:
            self.oc.put_file_contents(self.bucket_name + '/' + cache_file, json.dumps(resp))
        except Exception as e:
          log.warn('Unable to add object %s: %s', cache_file, repr(e))
          pass
        else:
          log.info('Added cache object %s', cache_file)

      return resp

  def get(self, cache_file):
    if self.enabled:
      log.debug('Looking for object %s', cache_file)

      rv = None
      cache_obj = None
      try:
        if self.s3:
          data = io.BytesIO()
          self.s3.meta.client.download_fileobj(self.bucket_name, cache_file, data)
          cache_obj = data.getvalue()
        elif self.oc:
          cache_obj = self.oc.get_file_contents(self.bucket_name + '/' + cache_file)
      except Exception as e:
        log.warn('Unable to load object %s: %s', cache_file, repr(e))
        pass

      if cache_obj:
        log.debug('Parsing object')
        try:
          rv = json.loads(cache_obj.decode("utf-8"))
        except Exception as e:
          log.warn('Unable to parse object %s: %s', cache_file, repr(e))
          pass
        else:
          log.info('Retrieved object %s', cache_file)

    return rv
