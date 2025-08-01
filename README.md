[![Binder](https://mybinder.org/badge_logo.svg)](https://mybinder.org/v2/gh/fomightez/dataframe2summary/master?filepath=index.ipynb)
# dataframe2summary
Repo for demonstrating scripts that convert dataframes / data tables into summarized dataframes / data tables. (Currently, the scripts demonstrated are located in [this repo](https://github.com/fomightez/text_mining).)

Click on `launch binder` badge above to spin up a sesion where you can step through the demos.

*These scripts take dataframes or tabular text (tables as text) as input.* 

If you have data as a table from elsewhere you can convert it/export into tabular text as tab-separated or comma-separated form and that can be used as input by any of the approaches here.

(Big picture, the main effort is taking something along the lines of [tidy data](https://r4ds.hadley.nz/data-tidy.html#sec-tidy-data) and making it a human-readable summary. In most cases, here this is generating as output a Pandas dataframe that displays nicely in Jupyter. Now, non-tidy though. I plan to add another, supplemental script in a similar vein. That script will take a tidy dataframe and uses openpyxl to produce a nice summary of a few related rows as 'block summaries' in an Excel spreadsheet.)

-----

## Demonstration notebooks

The intent is that there be (at least) **two** notebooks:  
The **first** notebook that opens in the active session demonstrates a script that makes it easy to convert a dataframe with groups and subgroups/states into a summary. For example, like the following from a dataframe.  
Examples of typical input and results (**the red annotation is just for illustration**):

![typical1](imgs/text_subgrp_example.png)  

![typical1](imgs/df_subgrp_example.png)  

The **second** notebook shows how to make a summary data table much like the first notebook produces; however, this script is specialized for binary data for the subgroups, i.e., they can only have two resulting states at most.  
Examples of typical input and results (**the red annotation is just for illustration**):

![data_table_binary](imgs/text_to_binary_first_example.png)  

![df_binary_summaries](imgs/df_based_binary_to_summaries.png)  


-----

Click on a `launch binder` badge on this page to spin up a sesion where you can make plots.

[![Binder](https://mybinder.org/badge_logo.svg)](https://mybinder.org/v2/gh/fomightez/dataframe2summary/master?filepath=index.ipynb)