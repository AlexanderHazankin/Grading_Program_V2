# Grading_Program_V2
This script is used to help a Teachers save time by grading and providing feedbacks for large classes by automating the process.

## Requirements
Requires Python 3.x to run.
openpyxl library: to work with Excel files
feedback_list module: to access predefined feedback list

## Usage
The script loads an Excel file named 'data.xlsx' which should be in the same directory.
The script accesses the active sheet in the workbook and iterates through the rows, starting from the second row.
The script expects the first column to be the student's name and the second column to be the student's score.
The script assigns the grade and feedback in the third and fourth columns respectively.
For each row, the script checks for the score and assigns a grade and feedback based on the predefined score ranges and feedback list.
If the score is missing, out of range or not a number, the script assigns an error message in the grade column.
The script saves the changes and closes the workbook.

## Credits
Created by Alexander Hazankin.

This code was created with the assistance of ChatGPT, a language model developed by OpenAI.

## Contact
For any questions or comments, you can reach me at:

https://www.linkedin.com/in/hazankin

https://github.com/AlexanderHazankin

https://replit.com/@Hazankin

## License
This project is licensed under the MIT [License](LICENSE).

Copyright (c) 2022 Alexander Hazankin.

Permission is hereby granted, free of charge.

## Note
The script uses the feedback_list module to access the predefined feedback list, make sure to import it if you move the script to another location.
If you want to use a different Excel file, change the file name in the script.
If you want to use different columns for the name, score and feedback, change the respective column_index_from_string value in the script.
If you want to change the score ranges or feedback list, change the grades and feedback_list variables in the script.



