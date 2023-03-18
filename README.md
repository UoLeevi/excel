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
      _ones_m;SEQUENCE(ROWS(X);;1;0);
      _ones_n;SEQUENCE(COLUMNS(X);;1;0);
      _w;IF(ROWS(w)=COLUMNS(X);w;SUM(y)/SUM(MMULT(X;_ones_n))*_ones_n);
      w_ols;_w-learning_rate*MMULT(TRANSPOSE((MMULT(X;_w)-y)*X);_ones_m)/ROWS(X);
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
    _ones_m;SEQUENCE(ROWS(X);;1;0);
    _ones_n;SEQUENCE(COLUMNS(X);;1;0);
    _y;y/y_max;
    _X;X/x_max;
    _w;SUM(_y)/SUM(MMULT(_X;_ones_n))*_ones_n;
    _iterations;IF(iterations=0;200;iterations);
    _learning_rate;IF(learning_rate=0;0.025;learning_rate);
    GRADIENT_DESCENT_BVLS;LAMBDA(w;_;
      LET(
        w_ols;w-_learning_rate*MMULT(TRANSPOSE((MMULT(_X;w)-_y)*_X);_ones_m)/ROWS(_X);
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
    y_max;MAXROW(ABS(y));
    x_max;MAXROW(ABS(X));
    _ones_m;SEQUENCE(ROWS(X);;1;0);
    _ones_n;SEQUENCE(COLUMNS(X);;1;0);
    _y;y/y_max;
    _X;X/x_max;
    scale_w;TRANSPOSE(x_max/y_max);
    scale_w_inv;TRANSPOSE(y_max/x_max);
    w;IF(ISOMITTED(w);SUM(_y)/SUM(MMULT(_X;_ones_n))*_ones_n;scale_w*w);
    m;SEQUENCE(ROWS(w);COLUMNS(w);0;0);
    v;SEQUENCE(ROWS(w);COLUMNS(w);0;0);
    state;HSTACK(HSTACK(m;v);w);
    ADAM;LAMBDA(state;t;
      LET(
        _m;INDEX(state;SEQUENCE(ROWS(state));1);
        _v;INDEX(state;SEQUENCE(ROWS(state));2);
        _w;INDEX(state;SEQUENCE(ROWS(state));3);
        g;MMULT(TRANSPOSE((MMULT(_X;_w)-_y)*_X);_ones_m);
        m;b_1*_m+(1-b_1)*g;
        v;b_2*_v+(1-b_2)*(g^2);
        _a;a*SQRT(1-b_2^t)/(1-b_1^t);
        g_adam;m/(SQRT(v)+e);
        w_ols;_w-_a*g_adam;
        w;IF(ISOMITTED(constraint_function);w_ols;scale_w*constraint_function(scale_w_inv*w_ols;scale_w_inv*g_adam;_a;scale_w_inv));
        CHOOSE({1,2,3};m;v;w)));
    result;REDUCE(state;SEQUENCE(iterations;;1);ADAM);
    scale_w_inv*INDEX(result;SEQUENCE(ROWS(result));3)))


# BOX_CONSTRAINT

=LAMBDA(lbound;ubound;
  LAMBDA(w;g;a;scale_w_inv;
    LET(
      w_lbounded;IF(w<lbound;lbound;w);
      IF(w_lbounded>ubound;ubound;w_lbounded))))

# LASSO_CONSTRAINT

=LAMBDA(reg;
  LAMBDA(w;g;a;scale_w_inv;
    LET(
      w_reg;ABS(w)-reg*a*scale_w_inv;
      IF(w_reg>0;w_reg;0)*SIGN(w))))

# COMBINE_CONSTRAINTS

=LAMBDA(constraint_1; constraint_2;
  LAMBDA(w;g;a;scale_w_inv;
    LET(
      w_1;constraint_1(w;g;a;scale_w_inv);
      w_2;constraint_2(w_1;g;a;scale_w_inv);
      w_2)))
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

#### Take & Drop

```
# TAKE (polyfill)

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


# DROP (polyfill)

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

table: m x n
colname: string
returns m x n_out

=LAMBDA(table;colname;
  INDEX(table;SEQUENCE(ROWS(table)-1)+1;MATCH(colname;INDEX(table;1;SEQUENCE(;COLUMNS(table)));0)))

```

#### Text functions

```
# TEXTCONTAINSANY

find_text: string | string[]
within_text: string
case_sensitive: boolean
returns: boolean

=LAMBDA(find_text;within_text;[case_sensitive];
  LET(
    case_sensitive;IF(ISOMITTED(case_sensitive);FALSE);
    OR(NOT(ISERROR(SEARCH("*"&find_text&"*";within_text))))))

```

#### Data manipulation

```
# TEXTSPLIT (polyfill)

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


# VSTACK (polyfill)

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


# HSTACK (polyfill)

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


# MAPROWS

A: m x n
function: 1 x n -> 1 x n_out
returns: m x n_out

=LAMBDA(A;function;
  LET(
    A_head;TAKE(A;1);
    A_tail;DROP(A;1);
    initial_value;RESHAPE(function(A_head);1);
    REDUCE(initial_value;A_tail;LAMBDA(B;a_row;
      VSTACK(B;RESHAPE(function(a_row);1))))))


# CARTESIANPRODUCT

=LAMBDA(a_1;[a_2];[a_3];[a_4];[a_5];[a_6];[a_7];[a_8];[a_9]; 
  LET(
    cartesian_prod;LAMBDA(set_1;[set_2];
      IF(ISOMITTED(set_2);
        set_1;
        LET(
          rows_1;ROWS(set_1);
          rows_2;ROWS(set_2);
          cols_1;COLUMNS(set_1);
          cols_2;COLUMNS(set_2);
          MAKEARRAY(rows_1*rows_2;cols_1+cols_2;
            LAMBDA(row;col;
              IF(col<=cols_1;
                INDEX(set_1;FLOOR.MATH((row-1)/rows_2)+1;col);
                INDEX(set_2;MOD(row-1;rows_2)+1;col-cols_1)))))));
    cartesian_prod(
      cartesian_prod(
        cartesian_prod(
          cartesian_prod(
            cartesian_prod(
              cartesian_prod(
                cartesian_prod(
                  cartesian_prod(a_1;a_2);a_3);a_4);a_5);a_6);a_7);a_8);a_9)))


# SPLIT.CARTESIANPRODUCT

=LAMBDA(delimiter;arg_1;[arg_2];[arg_3];[arg_4];[arg_5];[arg_6];[arg_7];[arg_8];[arg_9];
  LET(
    split;LAMBDA(x;TEXTSPLIT(x;;delimiter));
    n_args;IFS(
      ISOMITTED(arg_2);1;
      ISOMITTED(arg_3);2;
      ISOMITTED(arg_4);3;
      ISOMITTED(arg_5);4;
      ISOMITTED(arg_6);5;
      ISOMITTED(arg_7);6;
      ISOMITTED(arg_8);7;
      ISOMITTED(arg_9);8;
      TRUE;9);
    CHOOSE(n_args;
      split(arg_1);
      CARTESIANPRODUCT(split(arg_1);split(arg_2));
      CARTESIANPRODUCT(split(arg_1);split(arg_2);split(arg_3));
      CARTESIANPRODUCT(split(arg_1);split(arg_2);split(arg_3);split(arg_4));
      CARTESIANPRODUCT(split(arg_1);split(arg_2);split(arg_3);split(arg_4);split(arg_5));
      CARTESIANPRODUCT(split(arg_1);split(arg_2);split(arg_3);split(arg_4);split(arg_5);split(arg_6));
      CARTESIANPRODUCT(split(arg_1);split(arg_2);split(arg_3);split(arg_4);split(arg_5);split(arg_6);split(arg_7));
      CARTESIANPRODUCT(split(arg_1);split(arg_2);split(arg_3);split(arg_4);split(arg_5);split(arg_6);split(arg_7);split(arg_8));
      CARTESIANPRODUCT(split(arg_1);split(arg_2);split(arg_3);split(arg_4);split(arg_5);split(arg_6);split(arg_7);split(arg_8);split(arg_9)))))


# MAPARGS

=LAMBDA(args;func;
  LET(
    get;LAMBDA(arr;col;INDEX(arr;;col));
    n_args;COLUMNS(args);
    CHOOSE(n_args;
      MAP(get(args;1);func);
      MAP(get(args;1);get(args;2);func);
      MAP(get(args;1);get(args;2);get(args;3);func);
      MAP(get(args;1);get(args;2);get(args;3);get(args;4);func);
      MAP(get(args;1);get(args;2);get(args;3);get(args;4);get(args;5);func);
      MAP(get(args;1);get(args;2);get(args;3);get(args;4);get(args;5);get(args;6);func);
      MAP(get(args;1);get(args;2);get(args;3);get(args;4);get(args;5);get(args;6);get(args;7);func);
      MAP(get(args;1);get(args;2);get(args;3);get(args;4);get(args;5);get(args;6);get(args;7);get(args;8);func);
      MAP(get(args;1);get(args;2);get(args;3);get(args;4);get(args;5);get(args;6);get(args;7);get(args;8);get(args;9));func)))


# HIERARCHIZE
=LAMBDA(root;keys;parents;[sort_keys];[level];
  LET(
    sort_by_array;IF(ISOMITTED(sort_keys);keys;sort_keys);
    root_level;IF(ISOMITTED(level);0;level);
    keys_sorted;SORTBY(keys;sort_by_array);
    parents_sorted;SORTBY(parents;sort_by_array);
    children;FILTER(keys_sorted;parents_sorted=root;NA());
    is_leaf;ISNA(INDEX(children;1;1));
    root_record;HSTACK(root;root_level;is_leaf);
    get_descendants_with_levels;LAMBDA(result;child;VSTACK(result;HIERARCHIZE(child;keys;parents;sort_by_array;root_level+1)));
    IF(is_leaf;
      root_record;
      REDUCE(root_record;children;get_descendants_with_levels))))


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
      data;IF(COLUMNS(B)>1;BYROW(B;coltype);MAP(B;coltype));
      VSTACK(colname;data));

    SWITCH(coltype;
      "intercept";LET(
        data;SEQUENCE(ROWS(data_table)-1;;1;0);
        VSTACK(colname;data));

      "classification";LET(
        classifications;TRANSPOSE(DROP(SORT(UNIQUE(A));1));
        colnames;colname&"_"&classifications;
        data;N(B=classifications);
        VSTACK(colnames;data));

      "number+dummy";LET(
        colnames;CHOOSE({1,2};colname&"_dummy";colname&"_value");
        data;CHOOSE({1,2};N(ISNUMBER(--B));IF(ISNUMBER(--B);--B;0));
        VSTACK(colnames;data));

      "distance-from-max";LET(
        data;MAX(A)-B;
        VSTACK(colname;data));

      "distance-from-max+dummy";LET(
        colnames;CHOOSE({1,2};colname&"_dummy";colname&"_value");
        data;CHOOSE({1,2};N(ISNUMBER(--B));IF(ISNUMBER(--B);MAX(FILTER(--A;ISNUMBER(--A)))-B;0));
        VSTACK(colnames;data));

      NA())))))
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
    DROP(SCAN(seed;SEQUENCE(m+1);lcg_parkmiller);1)/(2^31-1)))


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
      X;DROP(s;;1);
      function(y;X))))))


```

### Miscellaneous

#### SEO related functions

```

# CLUSTER_KEYWORDS

=LAMBDA(keywords;urls;[match_count];
  LET(
    match_count;IF(ISOMITTED(match_count);4;match_count);

    \1;"Helper functions";
    get_row;LAMBDA(table;i;INDEX(table;i;SEQUENCE(;COLUMNS(table))));

    \2;"Clustering";
    uq_keywords;UNIQUE(keywords);
    uq_urls;UNIQUE(urls);
    len_keywords;LEN(uq_keywords);
    count_keywords;ROWS(uq_keywords);
    sorted_keywords;SORTBY(uq_keywords;len_keywords);
    seq;SEQUENCE(count_keywords);
    keyword_pair_common_url_count_matrix;COUNTIFS(urls;uq_urls;keywords;TRANSPOSE(sorted_keywords));
    keyword_pair_mask_matrix;MMULT(TRANSPOSE(keyword_pair_common_url_count_matrix);keyword_pair_common_url_count_matrix)>=match_count;
    keyword_match_count;MMULT(--keyword_pair_mask_matrix;SEQUENCE(COLUMNS(keyword_pair_mask_matrix);;1;0));
    keyword_pair_matrix;SORTBY(TRANSPOSE(IF(keyword_pair_mask_matrix;sorted_keywords;""));keyword_match_count;-1);
    clusters_matrix;merge_rows_with_overlap(keyword_pair_matrix);
    cluster_size;BYROW(clusters_matrix;LAMBDA(cluster;SUM(--(cluster<>""))));
    clusters;FILTER(clusters_matrix;cluster_size>1);
    no_clusters;FILTER(clusters_matrix;cluster_size=1);
    no_cluster_row;REDUCE(get_row(no_clusters;1);SEQUENCE(ROWS(no_clusters)-1;;2);LAMBDA(agg_row;i;LET(
      current_row;get_row(no_clusters;i);
      IF(current_row="";agg_row;current_row))));
    clusters_all;VSTACK(clusters;no_cluster_row);
    empty_column;INDEX("";SEQUENCE(ROWS(clusters_all)-1;;;0));
    extra_column_for_no_cluster;VSTACK(empty_column;"(no cluster)");
    result_matrix;HSTACK(extra_column_for_no_cluster;clusters_all);
    clusters_text;BYROW(result_matrix;LAMBDA(cluster;TEXTJOIN(", ";TRUE;cluster)));
    clusters_text))

# merge_rows_with_overlap (required helper function)

=LAMBDA(table;LET(
  get_row;LAMBDA(table;i;INDEX(table;i;SEQUENCE(;COLUMNS(table))));
  merge_row;LAMBDA(table;row_values;i;LET(
    cols;SEQUENCE(;COLUMNS(table));
    seq;SEQUENCE(ROWS(table));
    r;IF(row_values="";INDEX(table;i;cols);row_values);
    IF(seq=i;r;table)));
  initial_table;get_row(table;1);
  REDUCE(initial_table;SEQUENCE(ROWS(table)-1;;2);LAMBDA(agg_table;i;LET(
    current_row;get_row(table;i);
    seq_agg_table;SEQUENCE(ROWS(agg_table));
    overlapping_rows_mask;MAP(seq_agg_table;LAMBDA(agg_i;LET(
      agg_row;get_row(agg_table;agg_i);
      agg_row_x;IF(agg_row="";-1;agg_row);
      OR(agg_row_x=current_row))));
    overlap_exists;OR(overlapping_rows_mask);
    IF(overlap_exists;
      LET(
        first_overlapping_row_index;XMATCH(TRUE;overlapping_rows_mask);
        merge_row(agg_table;current_row;first_overlapping_row_index));
      VSTACK(agg_table;current_row))
  )))))

```
