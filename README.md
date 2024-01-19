
# Bluejay Delivery Software Developer Internship Assignment

## Introduction
This repository contains the solution for the Bluejay Delivery Software Developer Internship assignment. The assignment involved analyzing an Excel file containing timecard data to identify specific patterns related to employee shifts.

## Project Structure
- **src/com/rcv/rwxlsx/FinalRead.java**: Java source code file containing the solution.
- **data/Assignment_Timecard.xlsx**: Sample Excel file used for testing.

## How to Run
1. Ensure you have Java installed on your machine.
2. Open a terminal or command prompt.
3. Navigate to the project directory.
4. Compile and run the `FinalRead.java` file.

```bash
javac -cp ".;path/to/poi-<4.1.2>.jar" src/com/rcv/rwxlsx/FinalRead.java
java -cp ".;path/to/poi-<4.1.2>.jar" com.rcv.rwxlsx.FinalRead
```

Replace `path/to/poi-<4.1.2>.jar` with the actual path to the Apache POI library.

## Output
The program will print the names and positions of employees who meet the specified criteria:
- Worked for 7 consecutive days.
- Had less than 10 hours of time between shifts but greater than 1 hour.
- Worked for more than 14 hours in a single shift.

## Dependencies
- [Apache POI](https://poi.apache.org/): A Java library for reading and writing Microsoft Office files.

## Issues or Improvements
If you encounter any issues or have suggestions for improvements, please open an issue on this repository.

## Author
Guddu Kumar
## License
This project is licensed under the [MIT License](LICENSE).

