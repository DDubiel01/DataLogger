# Sniper Analytics V2.0

## Frames and their Uses

- **`rawdf`**: The raw data imported from Google  
  *(Columns: Sniper, Sniped, Date, Notes)*  

- **`names_dict`**: Dictionary to get plain names  
  *(Keys: Codenames, Values: Plain names)*  

- **`decodedf`**: The whitelist  
  *(Columns: Code, Name)*

- **`basicframe`**: Each Sniper's total snipes, total times sniped, and K/D  
  *(Columns: Snipes, Sniped, K/D)*

- **`comboframe`**: A big matrix of the number of times one has sniped another  
  *(Rows: snipers, Columns: people they have sniped)*

- **`truncdateframe`**: `rawdf` but the dates are in `MM-DD-YYYY` format

- **`dateframe`**: How many snipes were made on a given day  
  *(Index: date, Column: snipes — includes days with 0)*

- **`weekframe`**: `dateframe` reported per week instead of per day  
  *(Index: Date of the Sunday of the week, Column: snipes)*

- **`sn_dateframe`**: How many snipes each person had every day  
  *(Index: date, Columns: list of snipers)*

- **`sd_dateframe`**: Like `sn_dateframe`, but tracks how many times each was sniped  

- **`sn_byweek` / `sd_byweek`**: Same as above, resorted by **week**

- **`uniqueframe`**: How many different people a sniper has sniped  
  *(Index: sniper, Column: unique number of snipes)*

- **`firstframe`**: Date of first snipe/sniped and how many days they have recorded a snipe / been sniped

## Functions, Inputs, and Outputs

- **`import_data`**:  
  *(Input: credentials, book_name, data_sheet, whitelist_sheet; Output: rawdf, names_dict)*
- **`create_basicframe`**  
  *(Input: rawdf, names_dict; Output: basicframe)*
- **`create_comboframe`**  
  *(Input: rawdf, names_dict; Output: comboframe)*
- **`create_uniqueframe`**  
*(Input: comboframe, names_dict; Output: uniqueframe)*
- **`create_timetotal`**  
*(Input: rawdf; Output: dateframe, weekframe, truncdateframe, daterange)*
- **`create_timeind`**  
*(Input: names_dict, truncdateframe, daterange; Output: sn_ and sd_dateframe, sn_ and sd_byweek)*
- **`create_firstframe`**  
  *(Input: names_dict, sn_datframe, sd_dateframe; Output: firstframe)*

## Graphs and their Requirements

-**`weekdaypie`**: Pie chart of days of the week based on snipes on those days. *Requires: truncdateframe*

-**`overtime`**: Line chart of snipes over time *Requires: truncdateframe*

-**`topSn_overtime`**: Line chart of the top N snipers over time. *Requires: basicframe, sn_dateframe*

-**`topSn_pie`**: Pie chart of the top N snipers compared to the rest of the group. *Requires: basicframe, sn_dateframe*

-**`topSd_overtime`**: Line chart of the top N sniped over time. *Requires: basicframe, sd_dateframe*

-**`topSd_pie`**: Pie chart of the top N sniped compared to the rest of the group. *Requires: basicframe, sd_dateframe*
