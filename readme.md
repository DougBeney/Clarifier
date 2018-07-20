Clarifier: A Simple but Powerful Spreadsheet Filtering Language, Written in a CSV File
---

It's easy to get started with Clarifier.

You need the `clarifier.py` file in this repo and you also need to create a `config.csv` file.

See `config.example.csv` for an example config.csv file.

The first column of your `config.csv` file contains commands. The three commands that are availible are:

- FILE
- ROW
- COL

Note that there is no case-sensitivity, however uppercase is recommended.

The first thing you do when creating a `config.csv` file is add a `FILE` command to load your file.

It follows the following syntax:

```
FILE, input-file.csv, output-file.csv
```

Instead of saving, you can even print the file to your terminal using `~print~`.

```
FILE, input-file.csv, ~print~
```

After that, you can start filtering our rows and columns.

The `ROW` and `COL` commands take the same amount of parameters.

```
{ROW or COL}, {LOCATION}, {CONDITIONAL}, {VALUE}
```

Here is an example to filter column D for all instances of "CEO":

```
ROW, D, is, "CEO"
```

There are three main conditional options and four numerical conditional arguments:

- is --> value is equal to the cell value
- contains --> cell contains the value
- regcontains --> cell contains a regular expression value

You can inverse these three conditionals as follows:

- !is --> value is not equal to the cell value
- !contains --> cell does not contain the value
- !regcontains --> cell does not contain a regular expression value

The numerical conditionals are:

- >
- >=
- <
- >=

Especially when using numerical conditionals, you will want to tweak your start and end values for filtering.

Here are some examples:

```
COL, B:2, >, 5
```

This filters column B for values that are greater than 5. The first row is probably a label (Ex. "Player Score") so we use that colon to specify a startpoint. It will start filtering on **row 2**.

```
COL, B:2:10, >, 5
```

This does the same thing as the previous except we also define a stop point. It will stop filtering at row 10.


---

Also note that after you have entered a few lines for filtering, you can specify another `FILE`.

You can also support different file formats such as ODS, XLS, and more by installing various [PyExcel packages](http://docs.pyexcel.org/en/latest/).
