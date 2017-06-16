# European Option Pricing with C++/.Net

## Abstract

We will present few modeling theories that allow us to compute the premium of the currency options. Then, estimating the volatility term structure essential for computing the premiums.


In order to prevail from the risks, we will present trading strategies involving options. We will also, using and empiric study, justify our choices regarding the term structure models chosen through this work.


Finally, we will compute the premium using all the models and compare between each and every one of them, then highlight the use and importance of hedging exposure using Greeks.


Keywords : Forex, Pricing, Garman-Kohlhagen, Discrete Models, Jump Models, Hedging strategies, Volatility...


## Overview

This repository is part of a project to build a solution from scratch, to price and hedge Euro options on FX and present it in a UI.

The main tools are C++/.Net/C#.


## The main steps

### Data preprocessing
- Preprocess trains stops times and delays in right format.
- Preprocess stations data in right format.
- Create JSON file describing graph egdes between stations.

### Vizualization initialization
- Parse data.
- Create graphs.
- Preprocess trains' trips to find shortest path and extrapolate missing data about delays.
- Preprocess summary of trains' delays.
- Render interaction tools (sliders/buttons).
- Render initial map: stations, subsections.
- Render initial datatable.
- Render graph of delays over day.

### Rendering at each time change
- Compute active trips state.
- Compute network state.
- Cender trains.
- Cender subsections jams.
- cender datatable.

## Credits
The amazing work done by Michael Barry and Brian Card on the  [MBTA](http://mbtaviz.github.io/) has inspired me. Both for visual conception, and some tricky parts of code for geometrical calculations.

I also used the following javascript libraries: es6-shim, underscore, moment, d3, c3, jquery, bootstrap, datatables.

## Source code and raw data
Source code is available here on github.
Raw data comes from:
- Transilien gtfs files on their website
- Extraction of their API I made available on an AWS S3 container here.

