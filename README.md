# Datastream-toolbox

This repository contains sample python code for filtering Datastream global equity universe, automatically updating downloads using Datastream request tables, and converting retrieved data to long record-form tables.

McGill subscribes Datastream through Excel add-on (DFO). DFO has the following data download limits:

"No more than 300,000 values should be requested in a single request. There is also a timeout of 4 minutes for a single request that may be encountered if several datatypes are requested in a Static request for a large list. In these cases, requests should be broken down to return smaller sets of values."

Retrieving equity universe and downloading data from DFO is not as straightforward as Compustat Global that is available on WRDS. This toolbox is made publicly available to help anyone who only has Excel access to Datastream.

The data retrieval process consists of the following steps:

1. Retrieve static security information using WorldScope country lists: FTSE-ACWI-WSuniverse.xlsm for stocks from non-US FTSE-ACWI countries, and US-WSuniverse.xlsm for US stocks.

2. Apply filters to the stock universe. WSuniverse.py applies filters in Griffin, Kelly and Nardari (2010) and Chaieb, Langlois and Scaillet (2020)

3. Define filtered stock universe as user-defined lists (#L) of maximum 5,000 instruments each.

4. Download data by creating and querying a Request table. Datastream_automatic_update.py gives an example of downloading daily return data.


Contact me for questions or bugs: yiliu.lu@mail.mcgill.ca
