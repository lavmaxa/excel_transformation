The aim of this work was to match physical and visual layers of the spreadsheet. The algorithm achieved this aim by iterating through table cells and merging some of them according to proposed heuristics.  
The work was done with **Java, Apache POI** (*API for working with Excel documents*) and adapted software **FrameFinder** written with **Python**.  
As an input there was a **SAUS** dataset. After transforming the spreadsheets became more machine-readable for further **ETL**-processes.  
The average percent of correct transformations was **85.82%**


**This programm needs to be run via IDE with five arguments:**  
1. Directory to FrameFinder program  
2. Directory with input spreadsheets  
3. Directory with results of FrameFInder  
4. Directory for results of transformation and normalization  
5. Directory with spreadsheets corrected by an expert (for cheking percent of correct transformations)