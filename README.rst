gslides
===============================
.. image:: https://img.shields.io/badge/python-3.7-green.svg
  :target: https://www.python.org/downloads/release/python-370/
.. image:: https://img.shields.io/badge/code%20style_black-000000.svg
  :target: https://github.com/amvb/black


Full sphix documentation can be found `here <https://michael-gracie.github.io/gslides/>`_

Quick Installation
------------------

.. code-block:: bash

  pip install git+https://github.com/michael-gracie/gslides.git

Developer Instructions
----------------------

To install the package with development dependencies run the command

.. code-block:: bash

  pip install -e .[dev]

This will enable the following

- Unit testing using `pytest <https://docs.pytest.org/en/latest/>`_
  - Run ``pytest`` in root package directory
- Pre commit hooks ensuring codes style using `black <https://github.com/ambv/black>`_ and `isort <https://github.com/pre-commit/mirrors-isort>`_
- Sphinx documentation
  - To create sphinx run ``make html`` in package docs folder
  - To view locally run ``python -m http.server``
