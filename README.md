# ***Dataset Concatenation & Join Tool ReadMe***

## Table of Contents

[1.0 - Preface](#preface)

[2.0 - Overview](#overview)

[3.0 - Dictionary Selection](#dictionary-selection)

[4.0 - Data Importing](#data-import)

[5.0 - Data Join Options](#data-join)

[6.0 - Join Notes](#join-notes)

[7.0 - Data Export](#data-export)

# 

# <a name="preface"></a>1.0 - Preface

Due to GDPR rules regarding the test data this program was designed for, much of its original content has been heavily adjusted or removed outright. What follows is a stripped-down concept version, eliminating large sections of the transformation and standardisation process relating to personal information. 

The dummy data found here is representational of the original dataset formats, but has been entirely generated from scratch. All reference codes, test areas, values and comments are entirely randomised and have no basis in actual test data.

# <a name="overview"></a>2.0 - Overview

This program was originally designed as an aid for the concatenation and joining of multiple .xlsx test datasets to quickly flag up instances of missing join fields, inconsistent recording methods, and common formatting errors. The original datasets could broadly fit under four main categories:

-   Primary Datasets - First batch of test data, primarily numeric with basic participant details

-   Secondary Datasets - Second batch of test data, primarily numeric with basic participant details

-   Tertiary Datasets - Breakdowns of dialect data, generally very inconsistent across datasets, containing the majority of the participant info with a mixture of integers and strings, sometimes in the same data column

-   Supporting Datasets - Contains miscellaneous data for things like post-testing commentary

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

## 2.1 - Interpreter Settings

Generated in Python 3.8.0 using the Pycharm IDE with
the following interpreter settings:

  |***Package***     |***Version***  |
  |----------------- |---------------|
  |XlsxWriter        |3.0.3          |
  |Et-xmlfile        |1.1.0          |
  |Numpy             |1.23.1         |
  |Openpyxl          |3.0.10         |
  |Pandas            |1.4.3          |
  |Pip               |21.1.2         |
  |Python-dateutil   |2.8.2          |
  |Pytz              |2022.2.1       |
  |Setuptools        |57.0.0         |
  |Six               |1.16.0         |
  |Wheel             |0.36.2         |

#  <a name="dictionary-selection"></a>3.0 - Dictionary Selection

To filter out irrelevant data and standardise column headers for
concatenation, a dictionary/apply mapping file must first be generated.
There are three main formats that can be used for this process.

![alt text](https://github.com/dwrlewis/Site-Converter/blob/0ca3230f265415ba9d96eae3b9129f7832062c87/README%20Images/image1.png)

## 

## 3.1 - Default Dictionary

The default dictionary is included within the tool itself. It contains
relevant mappings for the randomised site data's key fields, including
reference ID's, timepoints, area of testing etc, and can be selected
using the checkbox at the top left of the user interface.

## 

## 3.2 - .xlsx & .txt Dictionary Import

Mappings can be imported using an .xlsx template with a list of columns
and their corresponding mappings. It is also possible to import mappings
from a .txt file, which uses standard python dictionary formatting.
Below is an example of the default dictionary in both .xlsx and .txt
formats.

# <a name="data-import"></a>4.0 - Data Importing

Data is categorised based on the dictionary remapping's. Each of the
four main data types have a set of fields that are unique to only that
type. For example, the presence of a column remapped as \'Resp. Score\'
would automatically map a file as primary data, whilst the presence of a
'Voc. Group' column would mark the file as a tertiary dataset. In
instances where both these fields are present, the former would take
priority due to order of primacy going from Primary \> Supporting.

![alt text](https://github.com/dwrlewis/Site-Converter/blob/0ca3230f265415ba9d96eae3b9129f7832062c87/README%20Images/image4.png)

## 

## 4.1 - Data Types

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

## 4.2 - Other Data

Data that is categorised in the other column experienced errors during
the import. This is generally a result of no viable mappings being
present in the file but can also be caused by import errors such as file
corruption. If the case of the latter, the file will be pasted to the
other list box with its imported error mapped onto the end of the
filename in the format "Filename.xlsx -- Error Description"

# <a name="data-join"></a>5.0 - Data Join Options

There are several options for how data can be consolidated, including
generating custom total fields, combining comment columns, and merging
duplicate fields across multiple datasets.

When data is imported, the number of columns per dataset is displayed.
If there are excessive numbers of columns in datasets marked as primary
or tertiary, then these will be flagged up in yellow, and a warning will
appear in the output box.

![alt text](https://github.com/dwrlewis/Site-Converter/blob/0ca3230f265415ba9d96eae3b9129f7832062c87/README%20Images/image5.png)

## 5.1 - Generating Primary Totals

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

## 5.2 - Generating Tertiary Totals

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

## 5.3 - Dropping Sub-Totals

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

## 5.4 - Consolidating Fields

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

# <a name="join-notes"></a>6.0 - Join Notes

After selections have been made and 'Join Data' has been pressed, an
output of the data join findings will be printed to the main list box.
This is split into three main sections:

## 

## 6.1 - Total Field Generation

This section will show the number of subtotal fields found, the new
fields generated, as well as note whether the original fields have been
dropped. If no corresponding fields were found this will also be
highlighted.

![alt text](https://github.com/dwrlewis/Site-Converter/blob/0ca3230f265415ba9d96eae3b9129f7832062c87/README%20Images/image6.png)

## 

## 6.2 - Dataset Joins

This section will display the results of the data joins. Given that this
tool is primarily intended for flagging data inconsistencies & errors,
outer joins are used exclusively to prevent any data loss. Each join
will flag up instances where new lines are added as a result of join
ID's that were not present in the primary data, but are in the
Secondary, Tertiary, or Supporting data.

![alt text](https://github.com/dwrlewis/Site-Converter/blob/0ca3230f265415ba9d96eae3b9129f7832062c87/README%20Images/image7.png)

[]{#_Toc111398055 .anchor}

## 6.3 - Field Consolidation

Outputs a list of how many fields of a particular type were found across
multiple datasets. For example, if the 'Area' column was found across
Primary, Tertiary and Supporting, it would be marked as '3 fields
consolidated'.

This section also includes the total number of lines & columns in the
final data, as well as instances where there are multiple Reference
Codes, either due to human input error or multiple test timepoints.

![alt text](https://github.com/dwrlewis/Site-Converter/blob/0ca3230f265415ba9d96eae3b9129f7832062c87/README%20Images/image8.png)

# <a name="data-export"></a>7.0 - Data Export

When exporting the data, it is possible to select either a new file path
or export the consolidated excel file directly to the original import
folder. The exported file is always saved as 'Site Data Joined.xlsx' and
will overwrite any file with this name in the output folder.

![alt text](https://github.com/dwrlewis/Site-Converter/blob/0ca3230f265415ba9d96eae3b9129f7832062c87/README%20Images/image9.png)

[]{#_Toc111398058 .anchor}

## 7.1 - Export Format

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
this.

![alt text](https://github.com/dwrlewis/Site-Converter/blob/0ca3230f265415ba9d96eae3b9129f7832062c87/README%20Images/image10.png)
