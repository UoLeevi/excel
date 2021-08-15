# Excel snippet collection

Collection of formulas and macros to use in Microsoft Excel.

## Formulas

#### Ordinary least squares

```
# OLS

y: n x 1
X: n x m
returns: m x 1

=LAMBDA(y;X;MMULT(MMULT(MINVERSE(MMULT(TRANSPOSE(X);X));TRANSPOSE(X));y))
```

#### Bounded variable least squares (implemented using gradient descent)
```
# BVLS - using recursion

y: n x 1
X: n x m
lbound: m x 1
ubound: m x 1
learning_rate: number, e.g. 0.00004
iterations: number, e.g. 50
w: m x 1
returns: m x 1

=LAMBDA(y;X;lbound;ubound;learning_rate;iterations;w;
  IF(iterations=0;w;
    LET(
      _ones_n;SEQUENCE(ROWS(X);;1;0);
      _ones_m;SEQUENCE(COLUMNS(X);;1;0);
      _w;IF(ROWS(w)=COLUMNS(X);w;SUM(y)/SUM(MMULT(X;_ones_m))*_ones_m);
      w_ols;_w-learning_rate*MMULT(TRANSPOSE((MMULT(X;_w)-y)*X);_ones_n)/ROWS(X);
      w_lbounded;IF(ROWS(lbound)=COLUMNS(X);IF(w_ols<lbound;lbound;w_ols);w_ols);
      w_bounded;IF(ROWS(ubound)=COLUMNS(X);IF(w_lbounded>ubound;ubound;w_lbounded);w_lbounded);
      BVLS(y;X;lbound;ubound;learning_rate;iterations-1;w_bounded))))


# BVLS - using loop

y: n x 1
X: n x m
lbound: m x 1
ubound: m x 1
iterations: number, e.g. 1000
learning_rate: number, e.g. 0.04
returns: m x 1

=LAMBDA(y;X;lbound;ubound;iterations;learning_rate;
  LET(
    x_max;MAXROW(X);
    y_max;MAXROW(y);
    _ones_n;SEQUENCE(ROWS(X);;1;0);
    _ones_m;SEQUENCE(COLUMNS(X);;1;0);
    _y;y/y_max;
    _X;X/x_max;
    _w;SUM(_y)/SUM(MMULT(_X;_ones_m))*_ones_m;
    _iterations;IF(iterations=0;200;iterations);
    _learning_rate;IF(learning_rate=0;0.025;learning_rate);
    GRADIENT_DESCENT_BVLS;LAMBDA(w;_;
      LET(
        w_ols;w-_learning_rate*MMULT(TRANSPOSE((MMULT(_X;w)-_y)*_X);_ones_n)/ROWS(_X);
        w_lbounded;IF(ROWS(lbound)=COLUMNS(_X);IF(w_ols<lbound;lbound;w_ols);w_ols);
        w_bounded;IF(ROWS(ubound)=COLUMNS(_X);IF(w_lbounded>ubound;ubound;w_lbounded);w_lbounded);
        w_bounded));
    TRANSPOSE(y_max/x_max)*REDUCE(_w;SEQUENCE(_iterations);GRADIENT_DESCENT_BVLS)))
```

#### Aggregation functions

```
# SUMROW

A: n x m
returns: 1 x m

=LAMBDA(A;MMULT(SEQUENCE(;ROWS(A);1;0);A))


# SUMCOLUMN

A: n x m
returns: n x 1

=LAMBDA(A;MMULT(A;SEQUENCE(COLUMNS(A);;1;0)))


# MAXROW

A: n x m
returns: 1 x m

=LAMBDA(A;MAP(SEQUENCE(;COLUMNS(A));LAMBDA(i;MAX(INDEX(A;;i)))))


# MAXCOLUMN

A: n x m
returns: n x 1

=LAMBDA(A;MAP(SEQUENCE(ROWS(A));LAMBDA(i;MAX(INDEX(A;i)))))


# MINROW

A: n x m
returns: 1 x m

=LAMBDA(A;MAP(SEQUENCE(;COLUMNS(A));LAMBDA(i;MIN(INDEX(A;;i)))))


# MINCOLUMN

A: n x m
returns: n x 1

=LAMBDA(A;MAP(SEQUENCE(ROWS(A));LAMBDA(i;MIN(INDEX(A;i)))))
```

#### Take & Skip

```
# TAKE

array: n x m
n: number

=LAMBDA(array;n;FILTER(array;SEQUENCE(ROWS(array))<=n))


# SKIP

array: n x m
n: number

=LAMBDA(array;n;FILTER(array;SEQUENCE(ROWS(array))>n))
```

#### Lookup table column values
```
# CLOOKUP

table_array: range
column_name: string

=LAMBDA(table_array;column_name;INDEX(OFFSET(table_array;1;0;ROWS(table_array)-1;COLUMNS(table_array));;MATCH(column_name;OFFSET(table_array;0;0;1;COLUMNS(table_array));0)))
```


