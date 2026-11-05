/******************************************************************************
  STAT 3660SEF
  Hong Kong tourism recovery and hotel room rates

  What this file does, in one sentence:
  It builds one row per month, then prints six results.

  The data folder sits next to this program.
  This program writes the output folder.

  Before you run it, paste your own folder between the quotes below.
  That one line is what every infile uses. It is blank on purpose.
  It should be the tourism folder, the one that contains the data folder.
******************************************************************************/

/* validvarname=v7: keep column names simple, so Excel headers become names SAS can use.
   dlcreatedir: if the output folder does not exist yet, create it.
   nodate: do not stamp today's date on every page. */
options validvarname=v7 dlcreatedir nodate;

/* Graphs stay off until the two charts. Otherwise SAS draws extras we do not want. */
ods graphics off;

/* proj is a nickname for the project folder. &proj means "paste that folder here".
   Leave this blank in the repo. Each person types their own SAS folder here (see README). */
%let proj = ;

/* This line makes sure the output folder exists. The charts and one working file are written there. */
libname gp "&proj/output";


/* ----- 1. Remember the reopen date ----- */
/* The text file has two lines. The first 9 characters are a date, like 06FEB2023.
   The words start at character 11. truncover means a short line is not an error. */
data policy;
    infile "&proj/data/reopen_policy.dat" truncover;
    length event $80;                          /* room for 80 characters of description */
    input date date9. @11 event $char80.;      /* read the date, jump to column 11, read the words */
    format date date9.;                        /* show the date as 06FEB2023, not a raw number */
run;

/* _null_ means "do the work, but do not save a new table".
   Find the line about resumption of travel, and store that date under the name reopen.
   Later, &reopen is 6 February 2023. */
data _null_;
    set policy;
    if index(upcase(event), "RESUMPTION") then call symputx("reopen", date);
run;


/* ----- 2. Visitors. The sheet is sideways. Make it one row per place per month. ----- */
/* C&SD table 650-80001. Row 4 is the header. Rows 5-16 are places.
   The last three rows of the sheet are notes, so the range stops at row 19.
   getnames=yes uses row 4 as the column names. replace means overwrite the table if it exists. */
proc import datafile="&proj/data/arrivals_650_80001.xlsx"
            out=arrivals_raw dbms=xlsx replace;
    range="Arrivals$A4:JM19";
    getnames=yes;
run;

/* One Excel row is one place, and every month is its own column (m200401, m200402, ...).
   The hotel files are the other way around: one row per month.
   This step stacks the sheet so both files can be joined. */
data arrivals_long;
    length nationality $80 market_group $16;   /* set the text width before any value is stored */
    set arrivals_raw;
    nationality = strip(nationality_or_region); /* strip removes leftover spaces on the name */

    if nationality = "" then delete;           /* blank row: throw it away */
    if upcase(substr(nationality, 1, 4)) = "NOTE" then delete;  /* a note row, not a place */

    /* Sort each place into a basket. Total is everyone. It is not added into a basket. */
    nat = upcase(nationality);
    if nat = "TOTAL" then market_group = "Total";
    else if nat = "CHINESE MAINLAND" then market_group = "Mainland";
    else if nat in ("MACAO", "TAIWAN", "NORTH ASIA",
                    "SOUTH AND SOUTHEAST ASIA") then market_group = "Short-haul";
    else if nat in ("THE AMERICAS", "EUROPE",
                    "AUSTRALIA, NEW ZEALAND AND SOUTH PACIFIC",
                    "MIDDLE EAST", "AFRICA") then market_group = "Long-haul";
    else if nat = "NOT IDENTIFIED" then market_group = "Other";
    else market_group = "Unmatched";           /* if this ever appears, a name was missed */

    /* mos is every month column, from January 2004 through August 2026.
       The loop walks across those columns. Each cell becomes its own row. */
    array mos{*} m200401--m202608;
    do i = 1 to dim(mos);
        /* The column is named m200401. Drop the m, and read 200401 as year 2004, month 01. */
        month_dt = input(substr(vname(mos{i}), 2), yymmn6.);
        arrivals = mos{i};                     /* the visitor count in that cell */
        output;                                /* write this one place-month, then do the next column */
    end;

    keep nationality market_group month_dt arrivals;  /* drop the wide columns; we do not need them now */
    format month_dt monyy7.;                   /* show the month as AUG2021, not a raw date number */
run;

/* Sort by month, then by basket. The next steps need the rows in that order. */
proc sort data=arrivals_long;
    by month_dt market_group;
run;

/* Add up the places inside each basket, for each month.
   Total is left out here, so it is not added on top of the places it already includes.
   noprint: do not print this. It is a working table, not one of the six results. */
proc means data=arrivals_long noprint nway;
    where market_group ne "Total";
    class month_dt market_group;
    var arrivals;
    output out=by_group(drop=_type_ _freq_) sum=arrivals;
run;

/* The published Total, one number per month. This is the visitor count the models use. */
proc means data=arrivals_long noprint nway;
    where market_group = "Total";
    class month_dt;
    var arrivals;
    output out=by_total(drop=_type_ _freq_) sum=arrivals_total;
run;

/* by_group is tall: several rows per month, one basket each.
   This step makes it wide: one row per month, with a column for each basket.
   first.month_dt is the first basket in that month. last.month_dt is the last one.
   A missing basket stays blank (the dot), instead of being carried over from last month. */
data group_wide;
    set by_group;
    by month_dt;
    length mainland short_haul long_haul other 8;
    retain mainland short_haul long_haul other; /* hold the numbers until the month is finished */
    if first.month_dt then do;
        mainland = .; short_haul = .; long_haul = .; other = .;
    end;
    if market_group = "Mainland" then mainland = arrivals;
    else if market_group = "Short-haul" then short_haul = arrivals;
    else if market_group = "Long-haul" then long_haul = arrivals;
    else if market_group = "Other" then other = arrivals;
    if last.month_dt then output;              /* write the row only when all baskets for the month are in */
    keep month_dt mainland short_haul long_haul other;
run;


/* ----- 3. Hotels. The date in the file is 202108, not a real SAS date. ----- */
/* firstobs=2 skips the header row. dlm="," means the columns are separated by commas.
   dsd means two commas in a row are an empty value, not a skipped column. */
data occupancy;
    infile "&proj/data/hotel_occupancy_monthly.csv" dlm="," dsd firstobs=2 truncover;
    length ym_c $8;
    input ym_c $ occupancy;                    /* example: 202108 and 71 */
    /* Characters 1-4 are the year. Characters 5-6 are the month. Day is always the 1st. */
    month_dt = mdy(input(substr(ym_c, 5, 2), 2.), 1, input(substr(ym_c, 1, 4), 4.));
    format month_dt monyy7.;
    keep month_dt occupancy;                   /* occupancy is the percent of rooms filled */
run;

/* Same layout as the occupancy file. room_rate is the average room price in Hong Kong dollars. */
data room_rate;
    infile "&proj/data/hotel_room_rate_monthly.csv" dlm="," dsd firstobs=2 truncover;
    length ym_c $8;
    input ym_c $ room_rate;
    month_dt = mdy(input(substr(ym_c, 5, 2), 2.), 1, input(substr(ym_c, 1, 4), 4.));
    format month_dt monyy7.;
    keep month_dt room_rate;
run;

/* nodupkey: if a month appears twice, keep the first copy. One month must be one row. */
proc sort data=occupancy nodupkey; by month_dt; run;
proc sort data=room_rate nodupkey; by month_dt; run;


/* ----- 4. Weather. Daily rows in, one average per month out. ----- */
/* The Observatory file has three title rows, then the days, then a legend at the bottom.
   firstobs=4 skips the titles. encoding=utf-8 reads the Chinese-capable text file correctly.
   A day is kept only when the year is four digits and the temperature is a real number.
   That drops the legend, because the legend is words, not a year. */
data hko_daily;
    infile "&proj/data/hko_daily_temp.csv" encoding="utf-8"
           dlm="," dsd firstobs=4 truncover;
    length year_c $12 month_c $8 rawval $24;
    input year_c $ month_c $ day_c $ rawval $ completeness $;
    if not (length(strip(year_c)) = 4 and notdigit(strip(year_c)) = 0) then delete;
    temp_c = input(rawval, ?? 8.);             /* ?? means a bad value becomes missing, not an error */
    if missing(temp_c) then delete;
    month_dt = mdy(input(month_c, 2.), 1, input(year_c, 4.));  /* every day in August shares 1 Aug */
    keep month_dt temp_c;
run;

/* Average the days into one temperature for the month. Still not printed. */
proc means data=hko_daily noprint nway;
    class month_dt;
    var temp_c;
    output out=temp_month(drop=_type_ _freq_) mean=temp_c;
run;

/* Write those monthly averages to a plain text file, then read that file back.
   This is the fixed-width read the course asks for.
   Column 1-4 is the year. Column 6-7 is the month. Column 9-16 is the temperature.
   z2. writes month 1 as 01, not 1. 8.4 writes the temperature with 4 decimal places. */
data _null_;
    set temp_month;
    year = year(month_dt);
    month = month(month_dt);
    file "&proj/output/hko_temp.dat";
    put @1 year 4. @6 month z2. @9 temp_c 8.4;
run;

data temp_raw;
    infile "&proj/output/hko_temp.dat" truncover;
    input @1 year 4. @6 month 2. @9 temp_c 8.4;
    month_dt = mdy(month, 1, year);            /* turn year and month back into the 1st of that month */
    format month_dt monyy7.;
    keep month_dt temp_c;
run;

/* A merge only works if every table is sorted by the same key. The key is the month. */
proc sort data=temp_raw; by month_dt; run;
proc sort data=by_total; by month_dt; run;
proc sort data=group_wide; by month_dt; run;


/* ----- 5. One row per month, and only if every fact is there. ----- */
/* merge lines the five tables up by month.
   in=ina means "this month was in the visitor-total table". The other in= flags mean the same
   for the baskets, occupancy, room price, and temperature.
   A month is kept only when all five are present, and only from August 2021 on.
   August 2021 is when the hotel files start. They end in July 2026. That is the 60 months. */
data analysis;
    merge by_total(in=ina) group_wide(in=ing)
          occupancy(in=ino) room_rate(in=inr) temp_raw(in=int);
    by month_dt;
    if month_dt >= "01aug2021"d and ina and ing and ino and inr and int;

    /* Set the width of these two words BEFORE either word is stored.
       train is 5 letters. holdout is 7. Without this, SAS sizes the box from the first word
       and stores holdout as holdo. The test would then find no rows. */
    length period_c $16 role $8;

    year = year(month_dt);                     /* 2021, 2022, ... used for the year-by-year table */
    arrivals_m = arrivals_total / 1000000;     /* visitors in millions, so the model number stays small */

    /* The notice says 6 February 2023, but every row is dated the 1st of the month.
       So the first After month is 1 February 2023. The dashed line on the graph stays on the 6th. */
    reopen_month = mdy(month(&reopen), 1, year(&reopen));
    if month_dt >= reopen_month then period_c = "After reopen";
    else period_c = "Before reopen";

    /* Twelve on/off switches, one per calendar month.
       If the row is March, m3 is 1 and the other eleven are 0.
       January is the baseline, so the formulas later use m2-m12 and leave m1 out. */
    cal_month = month(month_dt);
    array md{12} m1-m12;
    do i = 1 to 12;
        md{i} = (cal_month = i);
    end;

    /* Learning months run through December 2024. January 2025 onward is the hidden test. */
    if month_dt <= "01dec2024"d then role = "train";
    else role = "holdout";

    drop i cal_month reopen_month;             /* working variables, not part of the saved row */
    format arrivals_total mainland short_haul long_haul comma12.
           room_rate comma8. temp_c 5.1 arrivals_m 8.2;
    /* These labels are the column headings on the printed tables. */
    label month_dt       = "Month"
          arrivals_total = "Visitor arrivals"
          arrivals_m     = "Arrivals (millions)"
          mainland       = "Mainland arrivals"
          short_haul     = "Short-haul arrivals"
          long_haul      = "Long-haul arrivals"
          occupancy      = "Hotel occupancy (%)"
          room_rate      = "Average room rate (HKD)"
          temp_c         = "Mean temperature (C)"
          period_c       = "Reopening period"
          role           = "Sample role";
run;

proc sort data=analysis; by month_dt; run;


/* ----- 6. Two guesses. Learn through December 2024. Score the months after that. ----- */
/* ods exclude all hides the long regression printout. It is not one of the six results.
   Output 6 prints the four numbers that matter. */
ods exclude all;

/* Keep only the learning months. role was set above. */
data train;
    set analysis;
    if role = "train";
run;

/* Guess A: the room price depends only on which month of the year it is.
   outest=estA saves the fitted numbers. plots=none stops extra graphs.
   PROC REG stays open until QUIT. */
proc reg data=train outest=estA plots=none;
    model room_rate = m2 m3 m4 m5 m6 m7 m8 m9 m10 m11 m12;
run;
quit;

/* Guess B: the same calendar, plus how full the hotels were, visitors in millions, and temperature. */
proc reg data=train outest=estB plots=none;
    model room_rate = m2 m3 m4 m5 m6 m7 m8 m9 m10 m11 m12
                      occupancy arrivals_m temp_c;
run;
quit;

/* estA has extra rows. PARMS is the row of fitted numbers.
   Copy them to names b0, b2, b3, ... so they do not collide with the month switches
   when they are joined back onto the data. b0 is the January baseline. */
data coefA;
    set estA;
    where _TYPE_ = "PARMS";
    b0 = Intercept;
    b2 = m2; b3 = m3; b4 = m4; b5 = m5; b6 = m6; b7 = m7;
    b8 = m8; b9 = m9; b10 = m10; b11 = m11; b12 = m12;
    keep b0 b2 b3 b4 b5 b6 b7 b8 b9 b10 b11 b12;
run;

/* Put those fitted numbers on every month, including the hidden ones, and do the sum by hand.
   For a March row, only m3 is 1, so the guess is b0 + b3.
   abs_err_a is how many dollars the guess was wrong, with the minus sign removed. */
data scoreA;
    if _n_ = 1 then set coefA;                 /* read the one row of fitted numbers, once */
    set analysis;
    pred_a = b0 + b2*m2 + b3*m3 + b4*m4 + b5*m5 + b6*m6 + b7*m7
                + b8*m8 + b9*m9 + b10*m10 + b11*m11 + b12*m12;
    abs_err_a = abs(room_rate - pred_a);
run;

/* Same copy for guess B, plus the three extra fitted numbers.
   They are renamed so occupancy the coefficient is not confused with occupancy the data. */
data coefB;
    set estB;
    where _TYPE_ = "PARMS";
    b0 = Intercept;
    b2 = m2; b3 = m3; b4 = m4; b5 = m5; b6 = m6; b7 = m7;
    b8 = m8; b9 = m9; b10 = m10; b11 = m11; b12 = m12;
    b_occ = occupancy;
    b_arr = arrivals_m;
    b_temp = temp_c;
    keep b0 b2 b3 b4 b5 b6 b7 b8 b9 b10 b11 b12 b_occ b_arr b_temp;
run;

data scoreB;
    if _n_ = 1 then set coefB;
    set analysis;
    pred_b = b0 + b2*m2 + b3*m3 + b4*m4 + b5*m5 + b6*m6 + b7*m7
                + b8*m8 + b9*m9 + b10*m10 + b11*m11 + b12*m12
                + b_occ*occupancy + b_arr*arrivals_m + b_temp*temp_c;
    abs_err_b = abs(room_rate - pred_b);
run;

/* Line the two guesses up by month, and keep only the hidden test months (January 2025 on). */
data hold_err;
    merge scoreA(keep=month_dt role pred_a abs_err_a)
          scoreB(keep=month_dt pred_b abs_err_b);
    by month_dt;
    if role = "holdout";
run;

/* Walk down the test months. n counts them. sum_a and sum_b add the dollar misses.
   On the last month, divide by n. That average miss is the number that decides the winner.
   A smaller miss is better. end=last is true only on the final row. */
data mae_both;
    set hold_err end=last;
    n + 1;
    sum_a + abs_err_a;
    sum_b + abs_err_b;
    if last then do;
        model = "A"; mae = sum_a / n; output;
        model = "B"; mae = sum_b / n; output;
    end;
    keep model n mae;
run;

/* The average room price on the learning months. Needed for R-square, not for the decision. */
proc means data=scoreA noprint;
    where role = "train";
    var room_rate;
    output out=mean_rate mean=mean_y;
run;

/* One row per learning month, with the real price and both guesses. */
data train_fit;
    if _n_ = 1 then set mean_rate(keep=mean_y);
    merge scoreA(keep=month_dt role room_rate pred_a)
          scoreB(keep=month_dt pred_b);
    by month_dt;
    if role = "train";
run;

/* R-square asks how close the guess was on the months it was allowed to see.
   sst is how far the real prices sit from their own average.
   sse is how far the real prices sit from the guess.
   1 means a perfect match on those learning months. 0 means the guess ignored them.
   This is not the test. Guess B scores high here because 2023 is inside the learning months. */
data rsquare;
    set train_fit end=last;
    n + 1;
    sst + (room_rate - mean_y)**2;
    sse_a + (room_rate - pred_a)**2;
    sse_b + (room_rate - pred_b)**2;
    if last then do;
        rsquare_a = 1 - sse_a / sst;
        rsquare_b = 1 - sse_b / sst;
        output;
    end;
    keep n rsquare_a rsquare_b;
run;

/* Turn printing back on. Everything below this line is allowed to show. */
ods exclude none;


/* ----- 7. The six things on screen. ----- */

/* The line chart needs the three baskets, but only for the 60 months that survived the merge.
   Divide by one million so the axis is "2.5" and not "2500000". */
data plot_arrivals;
    merge by_group(in=g) analysis(in=a keep=month_dt);
    by month_dt;
    if g and a;
    if market_group in ("Mainland", "Short-haul", "Long-haul");
    arrivals_m = arrivals / 1000000;
run;

/* Turn graphs on, and save the picture as a PNG in the output folder. */
ods graphics on / reset width=9in height=5.2in imagename="arrivals_by_group" imagefmt=png;
ods listing gpath="&proj/output" image_dpi=150;

/* Output 1. Three lines. group= draws one line per basket.
   refline draws the dashed line on 6 February 2023. */
proc sgplot data=plot_arrivals;
    series x=month_dt y=arrivals_m / group=market_group lineattrs=(thickness=2);
    refline &reopen / axis=x label="6 Feb 2023" lineattrs=(pattern=shortdash);
    xaxis label="Month" fitpolicy=rotatethin;
    yaxis label="Visitor arrivals (millions)";
    title "Output 1 of 6. Who came back";
    title2 "The dashed line is 6 February 2023.";
    format month_dt monyy7. arrivals_m 4.1;
    keylegend / title="";
run;

/* Output 2. One dot per month. Colour is Before or After. */
ods graphics / reset width=8in height=5.2in imagename="rate_vs_occupancy" imagefmt=png;

proc sgplot data=analysis;
    scatter x=occupancy y=room_rate / group=period_c
            markerattrs=(symbol=circlefilled size=9);
    xaxis label="Hotel occupancy (%)";
    yaxis label="Average achieved room rate (HKD)";
    title "Output 2 of 6. Room rate and how full the hotels were";
    title2 "Each point is one month.";
    keylegend / title="";
run;

ods graphics off;

/* Output 3. Average visitors, fullness, room price, and temperature inside each year.
   2021 has 5 months. 2026 has 7. The title says so, so those years are not read as full years.
   maxdec=1 prints one decimal place. */
proc means data=analysis n mean maxdec=1;
    class year;
    var arrivals_total occupancy room_rate temp_c;
    title "Output 3 of 6. Year by year";
    title2 "2021 is August to December. 2026 is January to July.";
    format arrivals_total comma12. room_rate comma8.;
run;

/* Four sorts of the same 60 months. descending means biggest first.
   Each sort is saved under its own name so the original table stays in month order. */
proc sort data=analysis out=by_rate_hi; by descending room_rate; run;
proc sort data=analysis out=by_rate_lo; by room_rate; run;
proc sort data=analysis out=by_occ_hi; by descending occupancy; run;
proc sort data=analysis out=by_occ_lo; by occupancy; run;

/* obs=4 keeps the first four rows of each sorted table.
   Stacked in this order, rows 1-4 are the dearest, 5-8 the cheapest,
   9-12 the fullest, 13-16 the emptiest. which is the label printed beside them. */
data extremes;
    length which $28;                          /* set the width before the labels are stored */
    set by_rate_hi(obs=4) by_rate_lo(obs=4)
        by_occ_hi(obs=4) by_occ_lo(obs=4);
    if _n_ <= 4 then which = "Highest room rate";
    else if _n_ <= 8 then which = "Lowest room rate";
    else if _n_ <= 12 then which = "Highest occupancy";
    else which = "Lowest occupancy";
run;

/* Output 4. label uses the column headings from the label statement above.
   noobs hides the row number. notsorted keeps the four groups in the order we stacked them. */
proc print data=extremes label noobs;
    by which notsorted;
    var month_dt room_rate occupancy arrivals_total period_c;
    title "Output 4 of 6. The extreme months";
    format month_dt monyy7. arrivals_total comma12. room_rate comma8.;
run;

/* Output 5. Compare the After months with the Before months, for price and for fullness.
   class period_c splits the rows into those two groups.
   After comes first in the alphabet, so the printed difference is After minus Before.
   The months are one stretch of time, so the p-value is not proof. */
proc ttest data=analysis;
    class period_c;
    var room_rate occupancy;
    title "Output 5 of 6. After reopening minus before reopening";
    title2 "Diff (1-2) is After minus Before. Months are not independent.";
run;

/* Turn the two average misses into two labelled rows for the last print. */
data answer_mae;
    length description $72;
    set mae_both;
    if model = "A" then description = "Holdout error, A: calendar month only (HKD)";
    else description = "Holdout error, B: month + occupancy + arrivals + temperature (HKD)";
    result = mae;
    keep description result;
run;

/* The rsquare table is one row with two numbers. This makes it two rows, same shape as the misses.
   output writes a row immediately, so both descriptions are kept. */
data answer_r;
    length description $72;
    set rsquare;
    description = "R-square, A, fit through December 2024";
    result = rsquare_a;
    output;
    description = "R-square, B, fit through December 2024";
    result = rsquare_b;
    output;
    keep description result;
run;

/* Stack the four rows: miss A, miss B, R-square A, R-square B. */
data answer;
    set answer_mae answer_r;
    label description = "Comparison" result = "Result";
run;

/* Output 6. The decision is the first two rows. The smaller miss wins.
   The R-square rows are the fit on the learning months, not the test. */
proc print data=answer label noobs;
    title "Output 6 of 6. Which description fits monthly room rates";
    title2 "Holdout error is the average miss, in HKD, from January 2025 onward.";
    format result 10.3;
run;
title;                                         /* clear the title so it does not leak onto a later run */
