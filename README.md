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
learning_rate: number, e.g. 0.00004
iterations: number, e.g. 1000
returns: m x 1

=LAMBDA(y;X;lbound;ubound;learning_rate;iterations;
  LET(
    _ones_n;SEQUENCE(ROWS(X);;1;0);
    _ones_m;SEQUENCE(COLUMNS(X);;1;0);
    _w;SUM(y)/SUM(MMULT(X;_ones_m))*_ones_m;
    GRADIENT_DESCENT_BVLS;LAMBDA(w;_;
      LET(
        w_ols;w-learning_rate*MMULT(TRANSPOSE((MMULT(X;w)-y)*X);_ones_n)/ROWS(X);
        w_lbounded;IF(ROWS(lbound)=COLUMNS(X);IF(w_ols<lbound;lbound;w_ols);w_ols);
        w_bounded;IF(ROWS(ubound)=COLUMNS(X);IF(w_lbounded>ubound;ubound;w_lbounded);w_lbounded);
        w_bounded));
    REDUCE(_w;SEQUENCE(iterations);GRADIENT_DESCENT_BVLS)))
```
