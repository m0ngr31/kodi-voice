#!/usr/bin/env python

import json
import requests
import boto3
import botocore
import hashlib
import io
import logging

log = logging.getLogger(__name__)


class KodiCache():
  def __init__(self, aws_access_key_id=None, aws_secret_access_key=None, bucket_name=None):
    self.cache_enabled = False
    self.s3 = None

    log.info('Initalizing')
    if aws_secret_access_key and aws_access_key_id and bucket_name:
      self.s3 = boto3.resource('s3', aws_secret_access_key=aws_secret_access_key, aws_access_key_id=aws_access_key_id)

      self.bucket_name = bucket_name
      log.info('Accessing bucket %s', self.bucket_name)
      self.bucket = self.s3.Bucket(self.bucket_name)

      try:
        self.s3.meta.client.head_bucket(Bucket=self.bucket_name)
      except botocore.exceptions.ClientError as e:
        log.error('Error %s accessing bucket %s', e.response['Error']['Code'], self.bucket_name)
        # continue on without the cache
        pass
      else:
        self.cache_enabled = True

        if log.getEffectiveLevel() == logging.DEBUG:
          objs = self.ls()
          if not len(objs):
            log.debug('Object cache empty')
          else:
            log.debug('Objects in cache:')
            for o in objs:
              log.debug('  %s', o)

    if self.cache_enabled:
      log.info('Initialized')
    else:
      log.info('Disabled')


  def ls(self):
    if self.cache_enabled:
      return [key.key for key in self.bucket.objects.all()]
    else:
      return []


  def clear(self):
    if self.cache_enabled:
      log.debug('Clearing cache objects')
      for key in self.bucket.objects.all():
        key.delete()
      #bucket.delete()
      log.info('Cleared all cache objects')


  def add(self, cache_file, url, auth, command, timeout, wait_resp=True):
    try:
      r = requests.post(url, data=command, auth=auth, timeout=timeout)
    except requests.exceptions.ReadTimeout:
      if not wait_resp:
        # caller doesn't care about the response anyway -- this is mostly for
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

      if self.cache_enabled and cache_file:
        log.debug('Adding object %s', cache_file)
        self.s3.Object(self.bucket_name, cache_file).put(Body=json.dumps(resp))
        log.info('Created/updated object %s', cache_file)

      return resp


  def get(self, cache_file):
    if self.cache_enabled:
      log.debug('Looking for object %s', cache_file)
      try:
        data = io.BytesIO()
        self.s3.meta.client.download_fileobj(self.bucket_name, cache_file, data)
        log.debug('Loading object %s', cache_file)
        r = json.loads(data.getvalue().decode("utf-8"))
        log.info('Retrieved object %s', cache_file)
        return r
      except:
        log.info('No object %s', cache_file)
        pass
