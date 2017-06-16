# European Option Pricing with C++/.Net
Under MIT License.

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
- Preprocess data : interest rates, FX rates, volatility, ...
- Scrapping routine results preprocessing : cleaning, regexes, standardization.
- Create JSON/csv files regrouping most of the data.

### Implemented models and routines (Both finished and under development)

#### Continuous Diffusions :

- Brownian Motion
- Geometric Brownian Motion
- CIR
- Square Bessel Process
- Ornstein Uhlenbeck process
- Time-integrated Ornstein Uhlenbeck process UD
- Levy Processes
- Jump Diffusions


#### Gamma process :
- Variance-gamma process
- Geometric Gamma process
- Step Processes


#### Renewal process
- Poisson process 

#### Volatility modelling :
- Implied volatility
- Historical volatility

#### Pricing models :
- Garman & Kohlhagen's model
- Merton's jump model
- Monte-Carlo simulations w/o variance reduction techniques
- Tree models

## Source code and raw data
Source code is available here on github.
Raw data comes from:
- Bloomberg.
- Web Scrapping.
- Internal data.
