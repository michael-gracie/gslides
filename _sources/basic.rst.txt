Basic Usage
=========================================

The following page will provide an understanding around the class structure for ``gslides``. To accompany this documentation, see an example of the usage of ``gslides`` in this `notebook <https://github.com/michael-gracie/gslides/blob/main/notebooks/usage.ipynb>`_.

Data Flow
----------------------------------------

Taking an object oriented approach, ``gslides`` provides classes that represent objects in either Google sheets or Google slides to enable the user flexibility in terms of how there are creating their pipeline and executing their dependencies.

The diagram provides a visual representation of the objects and methods available.

.. image:: img/uml.png

The modularization of the pipeline allows the user to create new spreadsheets, sheets and presentations or simply get existing sheets and presentations. Additionally, if the user wants to only use a portion of the pipeline they can choose so.

User supplied ID's
------------------------------------------

From existing spreadsheets and presentations it is possible to retrieve ``spreadsheet_id``, ``sheet_id``, ``presentation_id`` through simply looking at a URL.

.. image:: img/user_id.png

Initializing classes
------------------------------------------

The ``Spreadsheet``, ``Frame`` and ``Presentation`` objects are all initialized through either the ``create()`` or ``get()`` class method.

Passing ``Chart`` and ``Table`` objects to ``Presentation.add_slide()``
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

The ``Chart`` and ``Table`` objects are unique in the fact that they don't need to be created to be passed to the ``Presentation.add_slide()`` method. The layout and slide margin parameters set in the ``Presentation.add_slide()`` method will optimize the size of the ``Chart`` and ``Table`` objects that are to be created. The ``Presentation.add_slide()`` method handles both the creation of the objects, creation of the slide and the copy of the objects all in one.

.. note::
   There is no restriction on utilizing the ``create()`` method of the ``Chart`` and ``Table`` classes. For ``Chart``, the method will simply create a chart in Google sheets without optimizing the sizing of that chart for it's destination slide. For ``Table`` it will create a table in Google slides.
