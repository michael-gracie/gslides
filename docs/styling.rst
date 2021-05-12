Styling
=========================================

Setting font & palette
--------------------------------
It is possible to configure the font and palette at the package level. When one sets these configurations all future slides, chart and tables will inherit the set styling. To do so run the following

.. code-block:: python

  gslides.set_font('Arial')
  gslides.set_palette('black')


Available palettes
------------------------------

Check out `here <https://github.com/michael-gracie/gslides/blob/main/gslides/config/base_palettes.yaml>`_ for a ``yaml`` file containing the available palettes. If you wish to add a palette feel free to make a PR on the repo.

Using custom palettes
---------------------------------

Users can load custom palettes by creating the following file ``~/.gslides/custom_palettes.yaml``. The file should be formatted the same as `this <https://github.com/michael-gracie/gslides/blob/main/gslides/config/base_palettes.yaml>`_ config file. The colors can be listed by either `name <https://github.com/michael-gracie/gslides/blob/main/gslides/config/color_mapping.yaml>`_ or hex color code.
