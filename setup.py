from setuptools import setup
def parse_requirements(filename):
    """ load requirements from a pip requirements file """
    lineiter = (line.strip() for line in open(filename))
    return [line for line in lineiter if line and not line.startswith("#")]

setup(
  name = 'Kodi-Voice',
  packages = ['kodi_voice'],
  version = '1.1.4',
  description = 'A library for interfacing with Kodi with VUI platforms like Amazon Alexa, Google Home, and Cortana.',
  author = 'Joe Ipson',
  author_email = 'joe@ipson.me',
  url = 'https://github.com/m0ngr31/kodi-voice',
  zip_safe = False,
  include_package_data = True,
  keywords = ['kodi', 'voice', 'alexa'],
  classifiers = [],
  install_requires = ['requests', 'boto3', 'pyocclient', 'ConfigParser', 'num2words', 'roman', 'fuzzywuzzy']
)
