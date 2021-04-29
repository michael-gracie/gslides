gslides: Creating charts in Google slides
=========================================

``gslides`` is a Python package that helps analysts turn ``pandas`` dataframes into Google slides & sheets charts by configuring and executing Google API calls.

The package provides a set of classes that enable the user full control over the creation of new visualizations through configurable parameters while eliminating the complexity of working directly with the Google API.

Quick Installation
------------------

.. code-block:: bash

  pip install git+https://github.com/michael-gracie/gslides.git


Usage
------------------

Below is an example that only showcases a simple workflow. Full discussion around features can be found in the docs.

**1. Initialize package and connection**

.. code-block:: python

  import gslides
  from gslides import (
    CreatePresentation,
    CreateSpreadsheet,
    CreateFrame,
    Chart,
    Scatter,
    CreateSlide
  )
  from sklearn import datasets
  gslides.intialize_credentials(creds) #BringYourOwnCredentials


**2. Create a presentation**

.. code-block:: python

  prs = CreatePresentation(name = 'demo pres')
  prs.execute()


**3. Create a spreadsheet**

.. code-block:: python

  spreadsheet = CreateSpreadsheet(
    title = 'demo spreadsheet',
    sheet_name = 'demo sheet')
  spreadsheet.execute()

**4. Load the data to the spreadsheet**

.. code-block:: python

  iris = datasets.load_iris(as_frame = True)['data']
  frame = CreateFrame(df = iris,
              spreadsheet_id = spreadsheet.spreadsheet_id,
              sheet_id = spreadsheet.sheet_id,
              overwrite_data = True
              )
  frame.execute()

**5. Create a scatterplot**

.. code-block:: python

  sc = Scatter(y_columns = ['sepal width (cm)'])
  ch = Chart(
      data = frame.data,
      x_column = 'sepal length (cm)',
      series = [sc],
      title = f'Demo Chart',
      x_axis_label = 'Sepal Length',
      y_axis_label = 'Sepal Width',
      legend_position = 'NO_LEGEND',
  )

**6. Create a slide with the scatterplot**

.. code-block:: python

  sld = CreateSlide(
      presentation_id = prs.presentation_id,
      charts = [ch],
      layout = (1,1),
      title = "Investigation into Fischer's Iris dataset",
      notes = "Data from 1936"
  )
  sld.execute()

**7. Navigate to the presentation**

.. image:: img/usage.png

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
