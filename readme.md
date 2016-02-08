json
----

**required**

- table_name: the name of the table

- sheet_cols: array of col names that will be mapped to db fields.

    - Can be a multidimensional array (useful for non-atomic sheet-rows)

    - Cols prefixed with underline are fixed values. (e.g.: _FOO will be FOO)

- db_cols: array of db fields sheet_cols get mapped to

**optional**

- rows_to_skip: number of row to skip (default = 0)

- sheet_index: index of the sheet (default = 0)

- not_empty: if any of this fields is empty no row get inserted


examples
--------

Excel-Table

    A           |B              |C
    ----------------------------------------
    Type/Lang   |Objective-C    |Swift
    ----------------------------------------
    Integer     |NSInteger      |Int
    String      |NSString       |String
    Map         |NSDictionary   |Dictionary
    NULL        |NULL           |

**Example 1**

json

    {
        "table_name":"types",
        "rows_to_skip":1,
        "db_cols":["type", "objective_c", "swift"],
        "sheet_cols":["a", "b", "c"]
    }

result

    type        |objective_c    |swift
    ----------------------------------------
    Integer     |NSInteger      |Int
    String      |NSString       |String
    Map         |NSDictionary   |Dictionary
    NULL        |NULL           |

**Example 2**

    {
        "table_name":"types",
        "rows_to_skip":1,
        "db_cols":["type", "language", "name"],
        "sheet_cols":[
            ["a", "_objective_c", "b"],
            ["a", "_swift", "c"]
        ]
        "not_empty":["name"]
    }

result

    type        |language       |name
    ----------------------------------------
    Integer     |objective_c    |NSInteger
    Integer     |swift          |Int
    String      |objective_c    |NSString
    String      |swift          |String
    Map         |objective_c    |NSDictionary
    Map         |swift          |Dictionary
    NULL        |objective_c    |NULL