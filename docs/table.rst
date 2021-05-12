Tables
=========================================

With ``gslides`` it's possible to take either existing data in Google sheets or a pandas ``DataFrame`` and visualize it in Google slides.

Styling
-----------------------------------------

The package allows the user to style the header row background and the stub column background as well.

.. image:: img/table.png

To enable a background on the stub column and header row while setting the color find the following parameters below and their corresponding default configurations.

.. code-block::

  header = True
  stub = False
  header_background_color = "black"
  stub_background_color= "black"


Sizing
----------------------------------------

Sizing is particularly difficult with tables and one needs to be thoughtful about the number of rows and width of columns that will be created. If the Google API returns an error around the minimum row / column width you can either reduce amount of data you are attempting to capture in the table or change the ``font_size`` setting.
