"""
Set up
"""
from setuptools import setup, find_packages

"""
This MUST match the repo name
If the repo name has dashes name
here should use underscores
"""
name = 'microsoft_pac'
__version__ = None
exec(open(f"{name}/_version.py").read())

with open('requirements.txt') as f:
    requirements = f.read().splitlines()

setup(name=name, version=__version__,
      author="@adamtav",
      author_email="adam@tavnets.com",
      maintainer_email="adam@tavnets.com",
      packages=find_packages(),
      include_package_data=True,
      install_requires=requirements
      )  # pragma: no cover
