# Excel snippet collection

Collection of formulas and macros to use in Microsoft Excel.

## Formulas

#### Ordinary least squares

```
# OLS

y: m x 1
X: m x n
returns: m x 1

=LAMBDA(y;X;MMULT(MMULT(MINVERSE(MMULT(TRANSPOSE(X);X));TRANSPOSE(X));y))
```

#### Bounded variable least squares (implemented using gradient descent)

```
# BVLS - using recursion

y: m x 1
X: m x n
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

y: m x 1
X: m x n
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

#### Constrained least squares

```
# CLS

=LAMBDA(y;X;[constraint_function];[w];[iterations];[a];[b_1];[b_2];[e];
  LET(
    iterations;IF(ISOMITTED(iterations);1000;iterations);
    a;IF(ISOMITTED(a);0.001;a);
    b_1;IF(ISOMITTED(b_1);0.9;b_1);
    b_2;IF(ISOMITTED(b_2);0.999;b_2);
    e;IF(ISOMITTED(e);0.00000001;e);
    y_max;MAXROW(y);
    x_max;MAXROW(X);
    _ones_n;SEQUENCE(ROWS(X);;1;0);
    _ones_m;SEQUENCE(COLUMNS(X);;1;0);
    _y;y/y_max;
    _X;X/x_max;
    scale_w;TRANSPOSE(x_max/y_max);
    scale_w_inv;TRANSPOSE(y_max/x_max);
    w;IF(ISOMITTED(w);SUM(_y)/SUM(MMULT(_X;_ones_m))*_ones_m;scale_w*w);
    m;SEQUENCE(ROWS(w);COLUMNS(w);0;0);
    v;SEQUENCE(ROWS(w);COLUMNS(w);0;0);
    state;HSTACK(HSTACK(m;v);w);
    ADAM;LAMBDA(state;t;
      LET(
        _m;INDEX(state;SEQUENCE(ROWS(state));1);
        _v;INDEX(state;SEQUENCE(ROWS(state));2);
        _w;INDEX(state;SEQUENCE(ROWS(state));3);
        g;MMULT(TRANSPOSE((MMULT(_X;_w)-_y)*_X);_ones_n);
        m;b_1*_m+(1-b_1)*g;
        v;b_2*_v+(1-b_2)*(g^2);
        _a;a*SQRT(1-b_2^t)/(1-b_1^t);
        w_ols;_w-_a*m/(SQRT(v)+e);
        w;IF(ISOMITTED(constraint_function);w_ols;scale_w*constraint_function(scale_w_inv*w_ols;scale_w_inv*(w_ols-_w);scale_w_inv*g));
        CHOOSE({1,2,3};m;v;w)));
    result;REDUCE(state;SEQUENCE(iterations;;1);ADAM);
    scale_w_inv*INDEX(result;SEQUENCE(ROWS(result));3)))


# BOX_CONSTRAINT

=LAMBDA(lbound;ubound;
  LAMBDA(w;step;g;
    LET(
      w_lbounded;IF(w<lbound;lbound;w);
      IF(w_lbounded>ubound;ubound;w_lbounded))))

# LASSO_CONSTRAINT

=LAMBDA(reg;
  LAMBDA(w;step;g;
    LET(
      a;step/g*reg;
      IF(w>a;w-a;IF(w<-a;w+a;0)))))
```

#### Aggregation functions

```
# SUMROW

A: m x n
returns: 1 x n

=LAMBDA(A;MMULT(SEQUENCE(;ROWS(A);1;0);A))


# SUMCOLUMN

A: m x n
returns: m x 1

=LAMBDA(A;MMULT(A;SEQUENCE(COLUMNS(A);;1;0)))


# MAXROW

A: m x n
returns: 1 x n

=LAMBDA(A;MAP(SEQUENCE(;COLUMNS(A));LAMBDA(i;MAX(INDEX(A;;i)))))


# MAXCOLUMN

A: m x n
returns: m x 1

=LAMBDA(A;MAP(SEQUENCE(ROWS(A));LAMBDA(i;MAX(INDEX(A;i)))))


# MINROW

A: m x n
returns: 1 x n

=LAMBDA(A;MAP(SEQUENCE(;COLUMNS(A));LAMBDA(i;MIN(INDEX(A;;i)))))


# MINCOLUMN

A: m x n
returns: m x 1

=LAMBDA(A;MAP(SEQUENCE(ROWS(A));LAMBDA(i;MIN(INDEX(A;i)))))
```

#### Take & Skip

```
# TAKE

A: m x n
i: number
j: number
returns: i x j

=LAMBDA(A;[i];[j];
  LET(
    m;ROWS(A);
    n;COLUMNS(A);
    i;IF(ISOMITTED(i);m;i);
    j;IF(ISOMITTED(j);n;j);
    INDEX(A;SEQUENCE(i);SEQUENCE(;j))))


# SKIP

A: m x n
i: number
j: number
returns: m-i x n-j

=LAMBDA(A;[i];[j];
  LET(
    m;ROWS(A);
    n;COLUMNS(A);
    i;IF(ISOMITTED(i);0;i);
    j;IF(ISOMITTED(j);0;j);
    INDEX(A;SEQUENCE(m-i;;i+1);SEQUENCE(;n-j;j+1))))
```

#### Lookup table column values

```
# CLOOKUP

table: range
colname: string

=LAMBDA(table;colname;
  LET(
    colname;TRANSPOSE(TEXTSPLIT(colname;";"));
    INDEX(table;SEQUENCE(ROWS(table)-1)+1;MATCH(colname;INDEX(table;1;SEQUENCE(;COLUMNS(table)));0))))

```

#### Data manipulation

```
# TEXTSPLIT

string: string
separator: character
returns: m x 1

=LAMBDA(string;separator;
  LET(
    string;separator&string&separator;
    char_indexes;SEQUENCE(LEN(string));
    chars;MID(string;char_indexes;1);
    sep_indexes;FILTER(char_indexes;chars=separator);
    indexes;SEQUENCE(ROWS(sep_indexes)-1);
    start_nums;INDEX(sep_indexes;indexes)+1;
    num_chars;INDEX(sep_indexes;indexes+1)-start_nums;
    MID(string;start_nums;num_chars)))


# VSTACK

A: m_a x n_a
B: m_b x n_b
returns: m_a + m_b x max(n_a,n_b)

=LAMBDA(A;B;
  LET(
    m_a;ROWS(A);
    m_b;ROWS(B);
    n_a;COLUMNS(A);
    n_b;COLUMNS(B);
    n;MAX(n_a;n_b);
    i;SEQUENCE(m_a+m_b);
    j;SEQUENCE(;n);
    IF(i<=m_a;INDEX(A;i;j);INDEX(B;i-m_a;j))))


# HSTACK

A: m_a x n_a
B: m_b x n_b
returns: max(m_a,m_b) x n_a + n_b

=LAMBDA(A;B;
  LET(
    m_a;ROWS(A);
    m_b;ROWS(B);
    n_a;COLUMNS(A);
    n_b;COLUMNS(B);
    m;MAX(m_a;m_b);
    i;SEQUENCE(m);
    j;SEQUENCE(;n_a+n_b);
    IF(j<=n_a;INDEX(A;i;j);INDEX(B;i;j-n_a))))


# DICT

=LAMBDA(key1;value1;[d];
  LAMBDA(key;
    IF(key=key1;value1;d(key)))


# RESHAPE

A: m_a x n_a
m: number
returns: m x (m_a*n_a)/m

=LAMBDA(A;m;
  LET(
    m_a;ROWS(A);
    n_a;COLUMNS(A);
    n;(m_a*n_a)/m+N(MOD(m_a*n_a;m)>0);
    r;SEQUENCE(m);
    c;SEQUENCE(;n);
    i;(c-1)*m+r-1;
    INDEX(A;MOD(i;m_a)+1;i/m_a+1)))


# FLATTEN

A: m x n
returns: m*n x 1

=LAMBDA(A;
  LET(
    m_a;ROWS(A);
    n_a;COLUMNS(A);
    m;m_a*n_a;
    i;SEQUENCE(m;;0);
    INDEX(A;MOD(i;m_a)+1;i/m_a+1)))
```

#### Data preparation

```
# PREPARECOLS

=LAMBDA(table;colname;[coltype];[data_table];
  LET(
    data_table;IF(ISOMITTED(data_table);table;data_table);
    colname;TRANSPOSE(TEXTSPLIT(colname;";"));
    A;CLOOKUP(table;colname);
    B;CLOOKUP(data_table;colname);

    IF(ISOMITTED(coltype);
      VSTACK(colname;B);

    IF(ISNONTEXT(coltype);LET(
      data;BYROW(B;coltype);
      VSTACK(colname;data));

    IF(coltype="intercept";LET(
      data;SEQUENCE(ROWS(data_table)-1;;1;0);
      VSTACK(colname;data));

    IF(coltype="classification";LET(
      classifications;TRANSPOSE(SKIP(SORT(UNIQUE(A));1));
      colnames;colname&"_"&classifications;
      data;N(B=classifications);
      VSTACK(colnames;data));

    IF(coltype="number+dummy";LET(
      colnames;CHOOSE({1,2};colname&"_dummy";colname&"_value");
      data;CHOOSE({1,2};N(ISNUMBER(--B));IF(ISNUMBER(--B);--B;0));
      VSTACK(colnames;data));

    IF(coltype="distance-from-max";LET(
      data;MAX(A)-B;
      VSTACK(colname;data));

    IF(coltype="distance-from-max+dummy";LET(
      colnames;CHOOSE({1,2};colname&"_dummy";colname&"_value");
      data;CHOOSE({1,2};N(ISNUMBER(--B));IF(ISNUMBER(--B);MAX(FILTER(--A;ISNUMBER(--A)))-B;0));
      VSTACK(colnames;data));

    VSTACK(colname;B))))))))))
```

#### Resampling

```
# RANDOMARRAY

m: number
seed: number
returns: m x 1

=LAMBDA([m];[seed];
  LET(
    m;IF(ISOMITTED(m);1;m);
    seed;IF(ISOMITTED(seed);1234;seed);
    lcg_parkmiller;LAMBDA(seed;i;MOD(48271*seed;2^31-1));
    SKIP(SCAN(seed;SEQUENCE(m+1);lcg_parkmiller);1)/(2^31-1)))


# SAMPLE

A: m_a x n
m: number
replacement: logical
seed: number
returns: m x n

=LAMBDA(A;[m];[replacement];[seed];
  LET(
    m_a;ROWS(A);
    replacement;IF(ISOMITTED(replacement);IF(ISOMITTED(m);TRUE;FALSE);replacement);
    m;IF(ISOMITTED(m);m_a;m);
    seed;IF(ISOMITTED(seed);1234;seed);
    INDEX(A;IF(replacement;
      RANDOMARRAY(m;seed)*(m_a-1)+1;
      TAKE(SORTBY(SEQUENCE(m_a);RANDOMARRAY(m_a;seed));m));SEQUENCE(;COLUMNS(A)))))


# BOOTSTRAP

TODO: nested arrays are not supported

=LAMBDA(y;X;function;r;[seed];
  LET(
    y_X;HSTACK(y;X);
    seed;IF(ISOMITTED(seed);1234;seed);
    SCAN(0;SEQUENCE(r;;seed);LAMBDA(a;seed;LET(
      s;SAMPLE(y_X;;;seed);
      y;TAKE(s;;1);
      X;SKIP(s;;1);
      function(y;X))))))






```
