# ***Dataset Concatenation & Join Tool ReadMe***

## Table of Contents

[Preface:](#preface)

[Interpreter Settings](#interpreter-settings)

[Overview](#overview)

[Dictionary Selection:](#dictionary-selection)

[Default Dictionary](#default-dictionary)

[.xlsx & .txt Dictionary Import](#xlsx-.txt-dictionary-import)

[Data Importing:](#data-importing)

[Data Types](#data-types)

[Other Data](#other-data)

[Data Join Options:](#data-join-options)

[Generating Primary Totals](#generating-primary-totals)

[Generating Tertiary Totals](#generating-tertiary-totals)

[Dropping Sub-Totals](#dropping-sub-totals)

[Consolidating Fields](#section-5)

[Join Notes](#join-notes)

[Total Field Generation](#section-6)

[Dataset Joins](#section-7)

[Field Consolidation](#_Toc111398055)

[Data Export](#data-export)

[Export Format](#_Toc111398058)

# 

# Preface

Much of the original content of both this guide and the tool itself have
been anonymised due to GDPR rules regarding the original test data it
was designed for. What follows is a stripped-down version of the
original, eliminating entire sections involving personal info, and using
test data that has been manually generated and entirely randomised. No
Reference Codes, Area's, test values or comments present in the test
data provided correlate to any of the original data this tool was
intended for.

## Interpreter Settings

This program was generated in Python 3.8.0 using the Pycharm IDE with
the following interpreter settings:

  ***Package***     ***Version***
  ----------------- ---------------
  XlsxWriter        3.0.3
  Et-xmlfile        1.1.0
  Numpy             1.23.1
  Openpyxl          3.0.10
  Pandas            1.4.3
  Pip               21.1.2
  Python-dateutil   2.8.2
  Pytz              2022.2.1
  Setuptools        57.0.0
  Six               1.16.0
  Wheel             0.36.2

## Overview

This tool is intended as an aid for the concatenation and joining of
multiple .xlsx datasets to quickly flag up instances of missing join
ID's, inconsistent values between files, and common formatting errors.
The data extracts this was intended for can be categorised into four
main types, as follows:

-   Primary Datasets -- First batch of test data

-   Secondary Datasets -- Second batch of test data

-   Tertiary Datasets -- Breakdowns of dialect data

-   Supporting Datasets -- Contains miscellaneous data for things like
    post-testing commentary

Datasets may not be cleanly split between these four categories, some
.xlsx files may for example contain both their primary and tertiary data
in one or be missing certain columns from either. The reason for this
level of inconsistency is twofold:

1.  The source data comes from a multitude of test sites across various
    countries, so data recording methods can vary significantly

2.  Most data has been entered manually into .xlsx files, not extracted
    from a source system. As such, many errors are a result of
    inconsistent input, e.g., mixing strings, integers, and floats.

The output of this program is intended to flag up inconsistencies across
both primary -- supporting data for individual sites, as well as between
different sites recording methods. Given that some measurements may be
recorded as multiple subtotal fields, single total fields, and string
comments in single cells with value breakdowns, perfect joining is not
feasible, but it is designed to partially align these fields for ease of
manual alignment.

# Dictionary Selection

To filter out irrelevant data and standardise column headers for
concatenation, a dictionary/apply mapping file must first be generated.
There are three main formats that can be used for this process.

![Text Description automatically
generated](vertopal_f6b5fd26a71243b6ad864fe2b63e577d/media/image1.png){width="5.515748031496063in"
height="5.118110236220472in"}

## 

## Default Dictionary

The default dictionary is included within the tool itself. It contains
relevant mappings for the randomised site data's key fields, including
reference ID's, timepoints, area of testing etc, and can be selected
using the checkbox at the top left of the user interface.

## 

## .xlsx & .txt Dictionary Import

Mappings can be imported using an .xlsx template with a list of columns
and their corresponding mappings. It is also possible to import mappings
from a .txt file, which uses standard python dictionary formatting.
Below is an example of the default dictionary in both .xlsx and .txt
formats.

# Data Importing

Data is categorised based on the dictionary remapping's. Each of the
four main data types have a set of fields that are unique to only that
type. For example, the presence of a column remapped as \'Resp. Score\'
would automatically map a file as primary data, whilst the presence of a
'Voc. Group' column would mark the file as a tertiary dataset. In
instances where both these fields are present, the former would take
priority due to order of primacy going from Primary \> Supporting.

![Graphical user interface, text Description automatically
generated](vertopal_f6b5fd26a71243b6ad864fe2b63e577d/media/image4.png){width="5.511811023622047in"
height="5.118110236220472in"}

## 

## Data Types

At least one primary set of data is mandatory, if no files contain
primary data fields, the import will be cancelled. However, all other
datasets are optional providing there is at least one set of additional
data to join. For example, it is possible to join just secondary data
without any available tertiary or supporting.

If no viable fields are found in a file, it is reloaded with the next
row set as the column header. This is repeated 5 times to check if a
viable header line is present, after which the file is marked as
non-standard. The datasets being tested do not exceeds several hundred
lines in most instances, so this does not have a significant impact on
reload times.

## 

## Other Data

Data that is categorised in the other column experienced errors during
the import. This is generally a result of no viable mappings being
present in the file but can also be caused by import errors such as file
corruption. If the case of the latter, the file will be pasted to the
other list box with its imported error mapped onto the end of the
filename in the format "Filename.xlsx -- Error Description"

# Data Join Options

There are several options for how data can be consolidated, including
generating custom total fields, combining comment columns, and merging
duplicate fields across multiple datasets.

When data is imported, the number of columns per dataset is displayed.
If there are excessive numbers of columns in datasets marked as primary
or tertiary, then these will be flagged up in yellow, and a warning will
appear in the output box.

![](vertopal_f6b5fd26a71243b6ad864fe2b63e577d/media/image5.png){width="5.515748031496063in"
height="5.118110236220472in"}

## Generating Primary Totals

If the primary dataset is marked as having a large number of columns
(\>100), this is generally because the data contains individual test
values, rather than any categorical sub-totals or total fields. In these
instances, selecting "Primary Sub-Totals" will generate two new fields
from these values:

-   \'Resp. Score AT \' -- A raw total of all test values

-   \'Art. Score AT \' -- A total of all test values greater than 1

This is generally a straightforward combination due to the source data
being a system export without human input, so requires minimal if any
correction other than checking for completeness.

## Generating Tertiary Totals

Tertiary data recording methods are generally far more inconsistent
across test areas, so will generate a large volume of columns when
multiple tertiary datasets from different areas are loaded. Selecting
"Tertiary Sub-Totals" will generate up to five new fields depending on
the data present and dictionary mappings used:

-   \'Merged M/O Notes\' -- Consolidates all main language (M-L) & other
    language (O-L) Note fields into a single cell, marking the original
    column source, and text wrapping to a new line for each columns
    data.

-   \'M-L Test AT \' -- Generates a new total field based on the
    presence of 'M-L' and 'Test' in the column header. Subtotals first
    have a new temporary column generated that is purged of all
    non-numeric data and converted to Integer, if possible, which is
    then added to the \'M-L Test AT' field in a loop. The purged
    subtotal columns are then dropped leaving only the new auto totalled
    field and original sub-totals.

-   \'O-L Test AT (tot.)\' -- Same as above, but searches for all other
    language totals and combines them to generate a single other
    language total field of all additional language scores.

-   \'O-L Test AT (subs)\' -- Same as above but creates a single total
    field from the subtotals of all other language columns.

-   \'Voc. Group AT (subs)' & \'Voc. Group AT (tot.)\' -- Generates a
    binary 0/1 field from the \'O-L Test AT' fields to calculate if
    other language total scores are above or belove accepted margins for
    language fluency.

## Dropping Sub-Totals

There is also the option to drop all original fields that the new totals
have been generated from. This is generally recommended for Primary data
totalling, where subtotal columns can number in the hundreds, and their
consolidation can greatly aid in data readability. This is generally
also advised as this data is significantly less prone to inconsistent
integer/string entries due to the nature of the data recording method
used in the original test data.

For Tertiary data it is recommended to review the auto-total fields
against the original data manually due to the degree of inconsistency.
Whilst the purging of string characters generally allows for conversion
of most data types to Integers/Floats, some value fields may contain
entries like "100% Eng, 50% Fr, 10% Ger", which would generate
abnormally large auto total amounts. Reference to the original field
data and consolidated comment fields should be made before another
extract is performed that drops the original fields.

## 

## Consolidating Fields

It is also common for datasets to contain the same columns across
primary, secondary, tertiary, and supporting data. For example, the
'Timepoint (M)' field may be present in primary and tertiary, whilst the
'Area' field may be present in all four. When joining data, instances of
duplicate fields will have suffixes added corresponding to their data
origin, e.g. SCND, TERT, SUPP.

The Consolidate Fields option is used to generate a single merged column
from these fields, using a priority order of Primary first, if blank
then use the Secondary, then Tertiary, and so forth. Much like the Drop
Sub-Totals option, it is recommended that data first be output without
this option selected to check for consistency across datasets. Reviewing
the 'Timepoint (M)' fields may for example show that the primary data
has its test month marked as 16, whilst a Secondary may have its data
marked as 14.

# Join Notes

After selections have been made and 'Join Data' has been pressed, an
output of the data join findings will be printed to the main list box.
This is split into three main sections:

## 

## Total Field Generation

This section will show the number of subtotal fields found, the new
fields generated, as well as note whether the original fields have been
dropped. If no corresponding fields were found this will also be
highlighted.

![](vertopal_f6b5fd26a71243b6ad864fe2b63e577d/media/image6.png){width="5.515748031496063in"
height="5.118110236220472in"}

## 

## Dataset Joins

This section will display the results of the data joins. Given that this
tool is primarily intended for flagging data inconsistencies & errors,
outer joins are used exclusively to prevent any data loss. Each join
will flag up instances where new lines are added as a result of join
ID's that were not present in the primary data, but are in the
Secondary, Tertiary, or Supporting data.

![](vertopal_f6b5fd26a71243b6ad864fe2b63e577d/media/image7.png){width="5.5in"
height="5.118110236220472in"}

[]{#_Toc111398055 .anchor}

## Field Consolidation

Outputs a list of how many fields of a particular type were found across
multiple datasets. For example, if the 'Area' column was found across
Primary, Tertiary and Supporting, it would be marked as '3 fields
consolidated'.

This section also includes the total number of lines & columns in the
final data, as well as instances where there are multiple Reference
Codes, either due to human input error or multiple test timepoints.

![](vertopal_f6b5fd26a71243b6ad864fe2b63e577d/media/image8.png){width="5.480314960629921in"
height="5.118110236220472in"}

# Data Export

When exporting the data, it is possible to select either a new file path
or export the consolidated excel file directly to the original import
folder. The exported file is always saved as 'Site Data Joined.xlsx' and
will overwrite any file with this name in the output folder.

![](vertopal_f6b5fd26a71243b6ad864fe2b63e577d/media/image9.png){width="5.5in"
height="5.118110236220472in"}

[]{#_Toc111398058 .anchor}

## Export Format

The exported 'Site Data Joined.xlsx' file is designed to flag up
instances of inconsistencies across datasets through the following:

Column ordering by category, additional non-standard columns are moved
to the end

-   Instances of duplicate ID's are highlighted in dark red

-   All blank cells within the data are highlighted in light red

-   Increased border thickness and emboldened text for any custom
    generated fields

-   Colour scaling for primary, secondary, and tertiary data values
    relative to expected amounts

This enables for easier isolation of anomalies in the data recording
methodology across test areas and how to account for
this.![](vertopal_f6b5fd26a71243b6ad864fe2b63e577d/media/image10.png){width="5.905511811023622in"
height="2.7401574803149606in"}
