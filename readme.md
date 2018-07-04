CSVlang: Simple syntax for filtering spreadsheets
---

config.csv follows this format:

```csv
my-file.csv,new-file.csv
desired_column,operation,search_string
...
```

Real example:

```csv
people.csv,people-updated.csv
b,=,CEO
b,=,CFO
```

To execute this, run csvlang.py in the same directory as config.csv.

This example takes people.csv and specifies to output to a new file called people-updated.csv

It searches column b for CEO and CFO. When it finds matches, it deletes that row.
