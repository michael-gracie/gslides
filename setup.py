"""Setup file for the package, aliases:

.. code:: bash

    $ python setup.py lint
"""
import os
import sys

from setuptools import find_packages, setup

# Pylint configuration
PYLINTRC_FILE = "pylintrc"
PYLINT_MINIMUM_SCORE = "9.0"
IGNORE_FOLDER = "test"

##############
# CORE PACKAGE
##############

NAME = "gslides"
PACKAGES = find_packages()
META_PATH = os.path.join("gslides", "__init__.py")
KEYWORDS = [""]
AUTHOR = "Michael Gracie"
AUTHOR_EMAIL = "12mpggslides@gmail.com"
LICENSE = "MIT License"

PROJECT_URLS = {
    "Documentation": "https://github.com/pages/michael-gracie/gslides/build/html/index.html",
    "Bug Tracker": "https://michael-gracie.github.io/gslides//issues",
    "Source Code": "https://michael-gracie.github.io/gslides/",
}

CLASSIFIERS = [
    "Intended Audience :: Developers",
    "Natural Language :: English",
    "Operating System :: OS Independent",
    "Programming Language :: Python",
    "Programming Language :: Python :: 3.7",
]

with open("requirements.txt", "r") as f:
    INSTALL_REQUIRES = f.read().split("\n")

EXTRAS_REQUIRE = {
    "docs": ["sphinx", "sphinx_rtd_theme", "furo"],
    "test": ["coverage", "pytest", "pytest-cov", "pytest-pylint", "mypy"],
    "qa": [
        "pylint",
        "pre-commit",
        "black",
        "isort",
        "mypy",
        "check-manifest",
        "flake8",
    ],
}

EXTRAS_REQUIRE["dev"] = (
    EXTRAS_REQUIRE["test"] + EXTRAS_REQUIRE["docs"] + EXTRAS_REQUIRE["qa"]
)

HERE = os.path.abspath(os.path.dirname(__file__))

DESCRIPTION = "Wrapper around Google APIs to create charts in Google Slides with python"

with open(os.path.join(HERE, "README.rst"), encoding="utf-8") as file_open:
    LONG_DESCRIPTION = file_open.read()

############
# Installing
############


def version_check():
    if sys.version_info < (3, 6):
        sys.exit("Python 3.6 required.")


def install_pkg():
    """Setup for the package"""

    setup(
        name=NAME,
        version="0.1.1",
        description=DESCRIPTION,
        long_description=LONG_DESCRIPTION,
        long_description_content_type="text/x-rst",
        url="https://github.com/michael-gracie/gslides",
        project_urls=PROJECT_URLS,
        author=AUTHOR,
        author_email=AUTHOR_EMAIL,
        license=LICENSE,
        python_requires=">=3.7.0",
        packages=PACKAGES,
        install_requires=INSTALL_REQUIRES,
        classifiers=CLASSIFIERS,
        extras_require=EXTRAS_REQUIRE,
        include_package_data=True,
        zip_safe=False,
    )


if __name__ == "__main__":
    version_check()
    install_pkg()
