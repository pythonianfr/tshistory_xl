# TSHISTORY XL

## What

This is an Excel client for [tshistory][tshistory].

* If you want to install the Xlwings only client from scratch, follow 1)
* If you want to install the PyXLL client from scratch, follow 3)
* If you want to upgrade the Xlwings only to the PyXLL one follow 2) then 3)
* If you want to revert to the Xlwings from the PyXLL one follow 4) then 1)

## 1) Install Xlwings client (recommanded)

```sh
pip install tshistory_xl
pip install xlwings
```

Close Excel.
Make sure that xlwings is registered:

```sh
xlwings addin install
```

Then type:

```sh
$ tsh xl-addin install
```

In Excel you should see two new tabs: `xlwings` and `Saturn`.

[tshistory]: https://hg.sr.ht/~pythonian/tshistory
[supervision]: https://hg.sr.ht/~pythonian/tshistory_supervision
[formula]: https://hg.sr.ht/~pythonian/tshistory_formula

## 2) Uninstall Xlwings client

Just type

```sh
$ tsh xl-addin uninstall
```

## 3) Install PyXLL client
### Installing Refinery Excel add-in
Once your package in installed (`pip install tshistory_xl`) just type the following command:
`tsh xl-addin install --pyxll`

## Installing PyXLL

### Python package

[References][pyxll]

```sh
pip install pyxll
```

### Excel add-in

```sh
pyxll install
```

Respond to the interactive prompt.
In case of doubt, use the proposed default option.


* Note carefully where the files will be installed, especially
the  **pyxll.cfg** file. You can keep the default location.
* To activate the free 30 day trial when the installer asks
“Do you have a PyXLL license key?” enter “n”. This will install
PyXLL without a license key, activating the 30 day free trial
automatically. After the 30 days trial, you will have to buy
a subscription here:  https://www.pyxll.com/pricing.html

### Configuration of PyXLL

Edit the **pyxll.cfg** file

This file is highly documented, you might want to read it extensively if
you want to explore the PyXLL options.

Here are the section that must be completed:
```sh
[PYTHON]
executable =
```


make sure that it points on the executable python you want to use for this operation
(i.e. in the correct conda-env)


You must also specify the pythonpath which point to the tshistory_xl package.
You  are encouraged to keep the `examples` section that will produce a PyXLL tab in Excel
with a lot of usefull options (logs and reload button)

```sh
pythonpath = c:/<the-path-to the python package tshsitory_xl>
            ./examples
```
---
Note: to be sure to know where this package is installed, you can use the
interactive python console as follows:

```sh
> python

> import tshistory_xl
> tshistory_xl.__file__
```

You will get a path like this:


`C:\\Users\\<user-name>\\miniconda3\\Lib\\site-packages\\tshistory_xl\\tshistory_xl\\__init__`.py


Remove the __init__.py and past this string as the pythonpath

---
```sh
[PYXLL]
modules = …
          macros
          error_handler
```

Just add **macros** and **error_handler** to the list

```sh
error_handler = error_handler.error_handler
```

The rest of the file can be left unchanged


### Final touches:

* This refinery-client explicitly uses the macro feature of excel. Make sure that
your excel file allow it (i.e. has an .xlm suffix)
* An error can pop at the launch of the excel file with the text:
“Error importing ‘macros”: OpenSSL 3.0 legacy provider failed to load…”
To avoid such error, you have to define a new environment variable set at false:
```sh
CRYPTOGRAPHY_OPENSSL_NO_LEGACY = 0
```

In Excel you should see two new tabs: `PyXLL Example tab` and `Ts-Refinery`
(instead of Saturn).


## 4) Uninstall PyXLL client

Just type

```sh
$ tsh xl-addin uninstall --pyxll
```

## Base use

On a brand new Excel sheet, you need initially a tab named
`TSHISTORY_CFG`, which must contain:

* in A1: `webapi`
* in B1: http://uri-of-the-tshistory-instance

|| A  | B |
|-|:-: | :--: |
|1|webapi | http://uri-of-the-tshistory-instance|


Then, let's create the conditions to play with one series.

- *Push data (Save them server side)*

_Warning: with this addin one can easly push data in production
database. Be sure to only use series begining with "test" until you
completly master the process to not make a mess._


|| test.series.1  | test.series.2 |
|-|:-: | :--: |
|2020-09-01|1 | 6|
|2020-09-02|2 | 5|
|2020-09-03|3 | 4|
|2020-09-04|4 | 3|
|2020-09-05|5 | 2|
|2020-09-06|6 | 1|


In a **new sheet**, let's go to `B1`. There we type the name of a
series (e.g. `test.series.1`), same thing for`B2`.

From `A2` to `A7`, type timestamps e.g. "2020-9-1", ...,
"2020-9-6". Make sure Excel really understand those as dates.

Then you have to create a `name` (e.g. using the name manager
accessible from the `formula` tab) for the range `B2:C7`, whose name
is e.g. `rwc_test_zone` (it is crucial that we have a prefix like
`rwc_`. **Only the data must be included into the range of the
name. The margins (series name and date) must be adjacent to its
range.**

Finally in the `Saturn` tab, click on `Save Tab`

You can watch the result in base with the adapted url :
http://uri-of-the-tshistory-instance/tsinfo?name=test.series.1

- *Pull data (Get them, client side)*

On another sheet, you can recreate the previous step:

- build a name (with the correct suffix)

- write in the margin the names of the series and the date

Click on `Get Tab` to obtain the data in Excel.


## Configuration options

### Prefix name options

In the example, we use a `rwc_` prefix. Each caracter has its meaning
and can be omited.

* `r`: indicate that the zone can read data from the database. Could
  be omited if this excel zone is only used for manual entry

* `w`: allows to write in the database. Should be omited if the zone
  is only used for data consultation

* `c`: put some colors in the name range. Very useful: most of the
  errors that you will encounter are linked to the size of the range
  that does not fit the data.


#### Specific request options:

* `f`: will fill the trailing empty values with the last non-empty one

* If `r_` name would return such data:


|| test.series.3  |
|-|:-: |
|2020-09-01|1 |
|2020-09-02| |
|2020-09-03|3 |
|2020-09-04| |
|2020-09-05| |


* `rf_` name would return :


|| test.series.3  |
|-|:-: |
|2020-09-01|1 |
|2020-09-02| |
|2020-09-03|3 |
|2020-09-04|3 |
|2020-09-05|3 |


* `_month_`: when pushing data, a value defined on one date of the
  month will be extrapolated (daily) for the whole month


### Layout

The name range can be placed anywhere on the sheet, as long as the
margins (with series and dates) are placed adjacently.

One could play a little with the data layout by letting empty cells in
the margins:


|| test.series.1 ||test.series.2 |
|-|:-: |-| :--: |
|| |I can| |
|2020-09-01|1 |put| 6|
|2020-09-02|2 |any| 5|
|2020-09-03|3 |comment| 4|
|2020-09-04|4 |in | 3|
|2020-09-05|5 |this| 2|
||also||here|
|2020-09-06|6 |column| 1|


### Revisiting the past

All the series are versionned, which mean anyone can access to a
previous version of the series. There are two ways to access it:

#### Whole name

After the first pulling of data, the sheet should be decorated in the
upper left corner of the name range with a comment "ASOF".

A date (*recognized as such by excel*) placed in this corner will allow
to view the state of all series at this given time.

Finally in the `Saturn` tab, click on `Get All`, and see the values
coming.


|*ASOF date should be here*| test.series.3  |
|-|:-: |
|2020-09-01|1 |

Note that when such a date is given, **the data can not be pushed in
the database** (iow you cannot rewrite the past from the excel
client).


#### By column

Sometimes you want to be able to see side by side the same series at
different times. You can display such data with additionnal argument
`asof`in the upper margin


|| test.series.4  | test.series.4(asof = 2020-09-03)|
|-|:-: | :--: |
|2020-09-01|1 | 1|
|2020-09-02|2 | 2|
|2020-09-03|3 | 3|
|2020-09-04|4 | #N/A|
|2020-09-05|5 | #N/A|
|2020-09-06|6 | #N/A


*Note: date must be in ISO format YYYY-MM-DD*

As before, the series with the `asof` parameter will not be pushed
when pressing the `Save tab` or `Save all` button.


#### Model backtest

When backtesting a forecast model, one will need a `staircase`
request, i.e. a request where the selected value dates are linked to
the insertion dates. It allows to evaluate the validity of a model
given a forecast horizon. For this, one has to use in the upper left
corner a new keyword `asofdelta=<number>` where the number is the
forecast horizon in hours.


|asofdelta=24| test.series.3  |
|-|:-: |
|2020-09-01|1 |



### `Not A Number` handling

By default, when the data are missing at a given date, the
corresponding cell will be filled with `#N/A`.

This default behaviour can be altered with some more columns options
`(blank=empy)`, `(blank=prev)`, `(blank=3.14)`


|| test.5  | test.6(blank=empy) | test.7(blank=prev) | test.8(blank=3.14) | test.9(blank=0) |
|-|:-: | :--: |:--: |:--: |:--: |
|2020-09-01|1   |1|1|1   |1|
|2020-09-02|#N/A| |1|3.14|0|
|2020-09-03|3   |3|3|3   |3|
|2020-09-04|#N/A| |3|3.14|0|


### Resampling

The excel addin allows to resample the data when pulling them with the
option `(agg = <method>)` where the method can be `mean`, `sum`,`max`,
`min`.

The resampling algorithm uses the dates in the left margin as
intervals for the resampling wich will led to an empty cell at the
end.


|| test.series.10  |
|-|:-: |
|2020-09-01|1|
|2020-09-02|1|
|2020-09-03|1|
|2020-09-04|1|
|2020-09-05|1|



|| test.series.10(agg=sum)|*comments*
|-|:-: | :--: |
|2020-09-01|3 |*Sum from date >= 2020-09-01 and date < 2020-09-04*|
|2020-09-04|2|*Sum from date >= 2020-09-04 and date < 2020-09-07*|
|2020-09-07|| *No computation here *|


#### Notes on resampling

* The last cell will be empty, in any case

* The resampled data won't be pushed

* If such resampling is reoccuring, we strongly advise to define a new
  resampled series, *server side*, with the [formula][formula] system
  of tshistory


### Common pitfalls

* Most of your errors will come from a range name with an incorrect
  form. Check it thoroughly. Check that all your left margin are
  dates, and that the upper margin does not have the same series
  called twice (with an exception when the series are asociated with
  an `asof` option)

* The error returned by `xlwings` are quite a mouthful. However, most
  of the error that will raise will provide a usefull comment bury
  somewhere, provided by the developpers of this addin. Your eyes may
  bleed because of it, but the solution might be there.

* It is quite easy to push some data in the database, it is also very
  simple to prevent it (use the `rc_` prefix)

* If one change the value of series inserted by a different process,
  the rule of updating such data might surprise you at first
  look. More information [here][supervision]


# Testing

While the environment works well with recent versions of python and
pandas, it might not work under the most recent versions of Excel
(timedelta tests fail under V16.0). Keep that in mind whenever you're
launching pytest.

Some tests might not work especially if you haven't configured your
local Postgres instance, since `pytest_sa_pg` requires commands from
your local Postgres installation, like `initdb`.

```shell
# For Linux, add the following to your .profile
export PATH=$PATH:/usr/lib/postgresql/{version_number}/bin/
source ~/.profile
```

```powershell
# For Windows
[System.Environment]::SetEnvironmentVariable('path', $Env:Programfiles + ".\PostgreSQL\{version_number}\bin;" + [System.Environment]::GetEnvironmentVariable('path', "User"),"User")

# Or add the variables manually through your panel
```
