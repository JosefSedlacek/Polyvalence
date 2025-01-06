# Polyvalence
VBA macros from Polyvalence project. Is the ``polyvalence.bas`` file you can find macros, which are included in my excel application.

### What is this project about
A data tool that allows you to track the performance of workers. It is a system of classifying workers into several levels according to their experience and results. In the polyvalence report I have projected all the data we have from production about the workers. First of all, these are the feedback reports that I download with a macro from SAP. Next is the attendance data that I get from the database using SQL queries. Finally, information about the supervisor's evaluation of the workers - I created an Excel application for this purpose.

So it is a total of three tools:
1. Excel application - this is where the evaluation of workers takes place
2. Polyvalence overview in PowerBI - shows how the workers are doing
3. Excel for updating - automatically set up to download this data:
    * SAP workers reports
    * Workers attendance - from intern database
    * Appraisal table - from excel app

Thanks to the polyvalence system we have an overview of the experience and skills of the workers. We know who makes the most NOK, on which machine the most downtime is created, we can check for a specific product whether the selected worker has experience with it, how many hours he has spent on this product in the past, when he last produced this product, on which machine and what was his efficiency. And a lot of other information.
