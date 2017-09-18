from setuptools import setup
from pip.req import parse_requirements

setup(
  name = 'Kodi-Voice',
  packages = ['kodi_voice'],
  version = '1.0.0',
  description = 'A library for interfacing with Kodi with VUI platforms like Amazon Alexa, Google Home, and Cortana.',
  author = 'Joe Ipson',
  author_email = 'joe@ipson.me',
  url = 'https://github.com/m0ngr31/kodi-voice',
  zip_safe = False,
  include_package_data = True,
  keywords = ['kodi', 'voice', 'alexa'],
  classifiers = [],
  download_url = 'https://github.com/m0ngr31/kodi-voice/tarball/1.0.0',
  install_requires = ['requests', 'ConfigParser', 'num2words', 'roman', 'fuzzywuzzy']
)
