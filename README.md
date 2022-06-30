# VBA Challenge

## Background
Stakeholder interests are to perform an analysis on thousands of stocks over multiple years using Excel and VBA.

## Overview of Project

The project challenge is to apply code that organizes, calculates, and formats information. The code uses for loops and if-then statements to create arrays that store values. This allowed the user to loop information into categories and to further format numbers with different colors. VBA allows a user to write a complicated script that performs complex analyses, which cannot only be done on Excel.

![VBA_Challenge](VBA_Challenge.png)

### Purpose

The purpose of this challenge is to create a VBA macro that can trigger pop-ups and inputs, read and change cells values, and format cells, using for loops, nested for loops, and conditionals to direct logic flow. 

## Analysis and Challenges

The biggest challenge was debugging syntax errors or array errors, remembering my syntax, and diagnosing the problem with the code when running into errors. 

## Results

The initial code loops over the entire year's worth of stock data for each stock of interest. This code needed 12 different iterations over a whole year's data. The results of two years that are analyzed is shown below:

![VBA_Challenge_2017](VBA_Challenge_2017.png)

![VBA_Challenge_2018](VBA_Challenge_2018.png)

The refractored code looped over all the data for a given year only once while still saving all the values needed to update the table.

![VBA_Challenge_2017_refactored](VBA_Challenge_2017_refactored.png)

![VBA_Challenge_2018_refactored](VBA_Challenge_2018_refactored.png)

The percent decrease in runtime is calculated below as (2017 data):

Original time - new time = overall decrease
0.7265625 - 0.0859375 = 0.640625

Percent runtime reduction = (Overall decrease / original time) * 100 
Percent runtime reduction = (0.640625/0.7265625) * 100 
Percent runtime reduction = 88.17204

The result is 88.17% overall reduction in runtime while maintaining the same functionality. 


## Summary

- What are the advantages or disadvantages of refactoring code?

The advantages of refactoring code is reduced runtime with equal functionality. But it comes with the cost of refactoring the code, which takes some time to do. Refactoring code could increase errors in already working code. 

- How do these pros and cons apply to refactoring the original VBA script?

The results of refactoring a working VBA script decreases runtime by 88.17% while still achieving the same results. The code is now easier to read when refactored without the need for an additional nested for loop. 
