# TSHISTORY XL

## What

This is an Excel client for [tshistory][tshistory].

## Removing the old versions

```sh
pip uninstall xl_data_hub
```

## Installation

```sh
pip install tshistory_xl
pip install xlwings
```

Close Excel.

Make sure that xlwings is registered (`xlwings addin install`).

Then type:

```sh
$ tsh xl-addin install
```

In Excel you should see two new tabs: `xlwings` and `Saturn`.

[tshistory]: https://hg.sr.ht/~pythonian/tshistory

If you have to uninstall the old proprietary version, do

```sh
tsh xl-addin uninstall-any --name ZATURN.xlam
```

## Base use

On a brand new Excel sheet, you need initially a tab named
`_SATURN_CFG`, which must contain:

* in A1: `webapi`
* in B1: http://uri-of-the-tshistory-instance

Then, let's create the conditions to play with one series.

In a new sheet, let's go to `B1`. There we type the name of a series,
e.g. `test` (we assume it's a daily series with data for 2020).

From `A2` to `A4`, type timestamps e.g. "2020-9-1", ...,
"2020-9-3". Make sure Excel really understand those as dates.

Then you have to create a `name` (e.g. using the name manager
accessible from the `formula` tab) for the range `B2:B4`, whose name
is e.g. `rwc_test_zone` (it is crucial that we have a prefix like `rwc_`).

Finally in the `Saturn` tab, click on `Get All`, and see the values
coming.


