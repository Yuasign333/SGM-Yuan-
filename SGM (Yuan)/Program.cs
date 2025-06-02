using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

class Program
{
    static void Main()
    {
        // ----- File Paths -----

        string dataPath; // CSV file to store student records

        string reportsPath; // Directory to store report files

        // ----- Array Variables -----

        string[] lines; // Lines read from the CSV file for bulk import

        string[] records; // Lines read from the CSV file for exporting all students

        string[] allLines; // Lines read from the CSV file for searching individual student records

        string[] existingLines; // Lines read from existing data file for initialization

        string[] existingRecords; // Existing records for duplicate checking

        string[] SortLines; // Lines for sorting records

        string[] allRecords; // All records for sorting after bulk import

        // ----- List Variables -----

        List<string> studentIDs; // List to store student IDs

        List<string> studentFirstNames; // List to store student first names

        List<string> studentLastNames; // List to store student last names

        List<double> grade1List; // List to store Data Structures grades

        List<double> grade2List; // List to store Programming 2 grades

        List<double> grade3List; // List to store Math Application in IT grades

        // ----- String Variables -----

        string input; // User input for menu selection

        string id; // Student ID input

        string lastName; // Last name input

        string firstName; // First name input

        string date; // Current date for report file naming

        string directory; // Directory path for data file

        string reportFile; // Report file name for exporting all students

        string searchID; // Student ID for searching individual records

        string filePath; // Path to the CSV file for bulk import

        string record; // Individual student record string

        string studentFile; // Individual student report file path

        string line; // Current line being processed

        string g1, g2, g3; // Grade strings for parsing

        string section; // Section identifier for report file (e.g., A, B, C)

        string updatedLine; // Updated line for writing to the report file

        string sectionFolder; // Folder path for section-specific reports

        string sectionFile; // File path for section-specific reports

        // ----- Double Variables -----

        double grade1; // Data Structures grade input

        double grade2; // Programming 2 grade input

        double grade3; // Math Application in IT grade input

        double avg; // Average grade calculation

        // ----- Integer Variables -----

        int added; // Number of records added during bulk import

        int skipped; // Number of records skipped during bulk import due to invalid grades or format

        // ----- Boolean Variables -----

        bool found; // Flag to check if a student record was found during
                    // 
        bool endProgram; // Flag to control program exit

        bool ValidInput; // Flag to check if the input is valid

        bool validFullName; // Flag to check if full name is valid

        bool idExists; // Flag to check if student ID already exists

        bool nameExists; // Flag to check if student name already exists

        bool validGrade1, validGrade2, validGrade3; // Flags for grade validation

        bool idValid; // Flag for ID format validation

        bool nameValid; // Flag for name format validation

        bool duplicate; // Flag for duplicate checking

        bool g1Valid, g2Valid, g3Valid; // Flags for grade parsing validation

        // ----- HashSet Variables -----

        HashSet<string> existingStudents; // HashSet to track existing student IDs

        HashSet<string> currentBulkIDs; // HashSet to track IDs in current bulk operation

        HashSet<string> sectionsWritten; // HashSet to track sections already written in the report file

        // ----- String Array Variables -----

        string[] parts; // Array to hold split parts of a line
        string[] p; // Array to hold split parts for searching

        // Initialize variables

        dataPath = @"C:\Users\yuanm\source\repos\SGM (Yuan)\SGM (Yuan)\bin\Debug\Data\students_grades.csv";

        reportsPath = "Report/Report_";

        studentIDs = new List<string>(); studentFirstNames = new List<string>(); studentLastNames = new List<string>(); grade1List = new List<double>(); grade2List = new List<double>(); grade3List = new List<double>();

        lastName = ""; firstName = "";

        grade1 = 0; grade2 = 0; grade3 = 0; added = 0; skipped = 0;

        found = false; endProgram = false;

        existingStudents = new HashSet<string>();

        // Initialize and load existing student data
        try
        {
            if (File.Exists(dataPath)) // Check if the data file exists
            {
                existingLines = File.ReadAllLines(dataPath); // Read all lines from the existing data file

                foreach (string existingLine in existingLines) // Check each line in the existing data file
                {
                    parts = existingLine.Split(','); // Split the line by commas to get individual fields

                    if (parts.Length >= 3) // Ensure there are at least 3 parts (ID, Last Name, First Name)
                    {
                        // Add the student ID and names to the respective lists and trim whitespaces for cleaner look
                        studentIDs.Add(parts[0].Trim());
                        studentLastNames.Add(parts[1].Trim());
                        studentFirstNames.Add(parts[2].Trim());
                    }
                }
            }
        }
        catch (FileNotFoundException)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("Data file not found. A new file will be created when you add your first student record.");
            Console.ForegroundColor = ConsoleColor.White;
        }
        catch (UnauthorizedAccessException)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("Access denied: Unable to read the data file. Please check file permissions.");
            Console.ForegroundColor = ConsoleColor.White;
        }
        catch (IOException ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"Error reading data file: {ex.Message}");
            Console.ForegroundColor = ConsoleColor.White;
        }
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"Unexpected error while loading data: {ex.Message}");
            Console.ForegroundColor = ConsoleColor.White;
        }

        // Main program loop
        while (!endProgram) // Loop until the user chooses to exit
        {
            try
            {
                // Display main menu
                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.Clear();
                Console.WriteLine("=======================================");
                Console.WriteLine("         STUDENT GRADE MANAGER         ");
                Console.WriteLine("=======================================\n");

                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("1. Add New Student Record\n");
                Console.WriteLine("2. Add Bulk Records from File\n");
                Console.WriteLine("3. Export All Students\n");
                Console.WriteLine("4. Export Individual Student Grades\n");
                Console.WriteLine("5. Exit\n");

                Console.ForegroundColor = ConsoleColor.White;
                Console.Write("Enter your choice: ");
                input = Console.ReadLine();

                Console.Clear();

                switch (input)
                {
                    case "1":
                        try
                        {
                            ValidInput = false; // Reset validation flag

                            // Input and validate student ID
                            Console.Write("\nStudent ID (format e.g. IT1A005): ");
                            id = Console.ReadLine().Trim();

                            while (!ValidInput)
                            {
                      // Validate ID format: Course Acronym [0][1] must be "IT", Year Level [2] must be '1',  // Section [3] is a letter, Student Number [4-6] are digits
                                if (id.Length == 7 && id[0] == 'I' && id[1] == 'T' && id[2] == '1' &&char.IsLetter(id[3]) && char.IsDigit(id[4]) && char.IsDigit(id[5]) && char.IsDigit(id[6]))
                                {
                                    // Check if ID already exists in the file
                                    idExists = false;

                                    try
                                    {
                                        if (File.Exists(dataPath))
                                        {
                                            lines = File.ReadAllLines(dataPath);
                                            foreach (string existingLine in lines)
                                            {
                                                parts = existingLine.Split(',');
                                                if (parts.Length > 0 && parts[0] == id)
                                                {
                                                    idExists = true;
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                    catch (IOException ex)
                                    {
                                        Console.ForegroundColor = ConsoleColor.Red;
                                        Console.WriteLine($"Error checking for duplicate ID: {ex.Message}");
                                        Console.ForegroundColor = ConsoleColor.White;
                                        Console.WriteLine("\nPress Enter to continue...");
                                        Console.ReadLine();
                                        break;
                                    }

                                    if (idExists)
                                    {
                                        Console.ForegroundColor = ConsoleColor.Red;
                                        Console.WriteLine($"\nA student with ID {id} already exists. Cannot add duplicate record.");
                                        Console.ForegroundColor = ConsoleColor.White;
                                        Console.WriteLine("\nPress Enter to try again...");
                                        Console.ReadLine();
                                        Console.Clear();
                                        Console.Write("\nStudent ID (format e.g. IT1A005): ");
                                        id = Console.ReadLine();
                                    }
                                    else
                                    {
                                        ValidInput = true;
                                    }
                                }
                                else
                                {
                                    Console.ForegroundColor = ConsoleColor.Red;
                                    Console.WriteLine("\nInvalid Student ID format.\n" + "\nCorrect format: IT (Capital Letter) (fixed course) + 1 (fixed year) +");
                                    Console.WriteLine("\n1 capital letter (Section) + 3 digits (Student Number) Overall: (e.g. IT1A005)");
                                    Console.ForegroundColor = ConsoleColor.White;
                                    Console.WriteLine("\nPress Enter to try again...");
                                    id = Console.ReadLine();
                                    Console.Clear();
                                    Console.Write("\nStudent ID (format e.g. IT1A005): ");
                                    id = Console.ReadLine();
                                }
                            }

                            // Input and validate student names
                            validFullName = false;

                            while (!validFullName)
                            {
                                // Validate last name
                                ValidInput = false;
                                while (!ValidInput)
                                {
                                    Console.Clear();
                                    Console.Write("\nLast Name: ");
                                    lastName = Console.ReadLine().Trim();

                                    if (string.IsNullOrWhiteSpace(lastName) || lastName.Any(char.IsDigit) ||
                                        !char.IsUpper(lastName[0]) || lastName.Substring(1) != lastName.Substring(1).ToLower())
                                    {
                                        Console.ForegroundColor = ConsoleColor.Red;
                                        Console.WriteLine("\nLast Name cannot be empty or contain numbers.");
                                        Console.WriteLine("First letter must be capital, the rest lowercase.");
                                        Console.ForegroundColor = ConsoleColor.White;
                                        Console.WriteLine("\nPress Enter to try again...");
                                        Console.ReadLine();
                                    }
                                    else
                                    {
                                        ValidInput = true;
                                    }
                                }

                                // Validate first name
                                ValidInput = false;

                                while (!ValidInput)
                                {
                                    Console.Clear();
                                    Console.Write("\nFirst Name: ");
                                    firstName = Console.ReadLine().Trim();

                                    if (string.IsNullOrWhiteSpace(firstName) || firstName.Any(char.IsDigit) ||
                                        !char.IsUpper(firstName[0]) || firstName.Substring(1) != firstName.Substring(1).ToLower())
                                    {
                                        Console.ForegroundColor = ConsoleColor.Red;
                                        Console.WriteLine("\nFirst Name cannot be empty or contain numbers.");
                                        Console.WriteLine("First letter must be capital, the rest lowercase.");
                                        Console.ForegroundColor = ConsoleColor.White;
                                        Console.WriteLine("\nPress Enter to try again...");
                                        Console.ReadLine();
                                    }
                                    else
                                    {
                                        // Check for duplicate full name with different ID
                                        nameExists = false;
                                        for (int i = 0; i < studentFirstNames.Count; i++)
                                        {
                                            if (studentFirstNames[i].Equals(firstName, StringComparison.OrdinalIgnoreCase) &&
                                                studentLastNames[i].Equals(lastName, StringComparison.OrdinalIgnoreCase) &&
                                                studentIDs[i] != id)
                                            {
                                                nameExists = true;
                                                Console.ForegroundColor = ConsoleColor.Red;
                                                Console.WriteLine($"\nName already exists under student ID: {studentIDs[i]}.");
                                                Console.ForegroundColor = ConsoleColor.White;
                                                Console.WriteLine("\nPress Enter to try again...");
                                                Console.ReadLine();
                                                break; // Exit to retry from last name
                                            }
                                        }

                                        if (!nameExists)
                                        {
                                            validFullName = true; // Both names are valid and not a duplicate
                                            ValidInput = true;
                                        }
                                        else
                                        {
                                            break; // Go back to last name
                                        }
                                    }
                                }
                            }

                            // Add validated student to lists
                            studentLastNames.Add(lastName);
                            studentFirstNames.Add(firstName);
                            studentIDs.Add(id);

                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.WriteLine($"\nStudent added successfully: {lastName}, {firstName} [{id}]");
                            Console.ForegroundColor = ConsoleColor.White;
                            Console.WriteLine("\nPress Enter to continue...");
                            Console.ReadLine();

                            // Input and validate grades
                            ValidInput = false;

                            while (!ValidInput)
                            {
                                try
                                {
                                    Console.Clear();
                                    Console.ForegroundColor = ConsoleColor.Yellow;
                                    Console.WriteLine("\n =====Grading System=====");
                                    Console.WriteLine($"\nThe Grade of {id} - {lastName}, {firstName} will be recorded as follows:");

                                    Console.Write("\nData Structures Grade: ");
                                    validGrade1 = double.TryParse(Console.ReadLine(), out grade1);

                                    Console.Write("\nProgramming 2 Grade: ");
                                    validGrade2 = double.TryParse(Console.ReadLine(), out grade2);

                                    Console.Write("\nMath App in IT Grade: ");
                                    validGrade3 = double.TryParse(Console.ReadLine(), out grade3);

                                    if (validGrade1 && validGrade2 && validGrade3 && grade1 >= 70 && grade1 <= 100 &&
                                        grade2 >= 70 && grade2 <= 100 && grade3 >= 70 && grade3 <= 100)
                                    {
                                        ValidInput = true;
                                        record = $"{id.Trim()},{lastName.Trim()},{firstName.Trim()},{grade1},{grade2},{grade3}";

                                        try
                                        {
                                            // Ensure directory exists before writing
                                           directory = Path.GetDirectoryName(dataPath);

                                            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                                            {
                                                Directory.CreateDirectory(directory); // Create the directory if it doesn't exist
                                            } 

                                            // Write student record to file
                                            using (StreamWriter sw = new StreamWriter(dataPath, true))
                                            {
                                                sw.WriteLine(record);
                                            }

                                            // Extract section from ID (e.g., IT1A005 -> 'A')
                                            section = id.Substring(3, 1).ToUpper();

                                            // Build section folder and file path
                                             sectionFolder = Path.Combine("Sections", $"Section{section}");
                                             sectionFile = Path.Combine(sectionFolder, "grades.csv");

                                            // Ensure section folder exists
                                            if (!Directory.Exists(sectionFolder))
                                            {
                                                Directory.CreateDirectory(sectionFolder);
                                            }

                                            // Write the same student record to the section-specific CSV file
                                            using (StreamWriter sectionWriter = new StreamWriter(sectionFile, true))
                                            {
                                                sectionWriter.WriteLine(record);
                                            }

                                            // Sort the records after adding a new one
                                            SortLines = File.ReadAllLines(dataPath);
                                            Array.Sort(SortLines);
                                            File.WriteAllLines(dataPath, SortLines);

                                            Console.ForegroundColor = ConsoleColor.Green;
                                            Console.WriteLine("\nStudent record saved successfully!");
                                            Console.ForegroundColor = ConsoleColor.White;
                                        }
                                        catch (UnauthorizedAccessException)
                                        {
                                            Console.ForegroundColor = ConsoleColor.Red;
                                            Console.WriteLine("\nAccess denied: Unable to save student record. Please check file permissions.");
                                            Console.ForegroundColor = ConsoleColor.White;
                                        }
                                        catch (DirectoryNotFoundException)
                                        {
                                            Console.ForegroundColor = ConsoleColor.Red;
                                            Console.WriteLine("\nDirectory not found: Unable to create or access the data directory.");
                                            Console.ForegroundColor = ConsoleColor.White;
                                        }
                                        catch (PathTooLongException)
                                        {
                                            Console.ForegroundColor = ConsoleColor.Red;
                                            Console.WriteLine("\nFile path is too long. Please use a shorter path.");
                                            Console.ForegroundColor = ConsoleColor.White;
                                        }
                                        catch (IOException ex)
                                        {
                                            Console.ForegroundColor = ConsoleColor.Red;
                                            Console.WriteLine($"\nError saving student record: {ex.Message}");
                                            Console.ForegroundColor = ConsoleColor.White;
                                        }
                                    }
                                    else
                                    {
                                        Console.ForegroundColor = ConsoleColor.Red;
                                        Console.WriteLine("\nGrades must be between 70 to 100. Please re-enter.");
                                        Console.ForegroundColor = ConsoleColor.White;
                                        Console.WriteLine("\nPress Enter to try again...");
                                        Console.ReadLine();
                                    }
                                }
                                catch (FormatException)
                                {
                                    Console.ForegroundColor = ConsoleColor.Red;
                                    Console.WriteLine("\nInvalid input format. Please enter numeric values for grades.");
                                    Console.ForegroundColor = ConsoleColor.White;
                                    Console.WriteLine("\nPress Enter to try again...");
                                    Console.ReadLine();
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine($"\nUnexpected error while adding student: {ex.Message}");
                            Console.ForegroundColor = ConsoleColor.White;
                        }
                        break;

                    case "2": // Bulk import from CSV file
                        try
                        {

                            // Enter this in case 2: C:\Users\yuanm\source\repos\SGM (Yuan)\SGM (Yuan)\bin\Debug\Data\bulk.csv

                            Console.Write("\nEnter CSV file path: ");
                            filePath = Console.ReadLine();

                            if (!File.Exists(filePath))
                            {
                                Console.ForegroundColor = ConsoleColor.Red;
                                Console.WriteLine("\nFile not found. Please check the path and try again.");
                                Console.ForegroundColor = ConsoleColor.White;
                                break;
                            }

                            // Read existing records to check for duplicates
                            existingRecords = new string[0];
                            try
                            {
                                if (File.Exists(dataPath))
                                {
                                    existingRecords = File.ReadAllLines(dataPath);
                                }
                            }
                            catch (IOException ex)
                            {
                                Console.ForegroundColor = ConsoleColor.Red;
                                Console.WriteLine($"Error reading existing records: {ex.Message}");
                                Console.ForegroundColor = ConsoleColor.White;
                                break;
                            }

                            try
                            {
                                lines = File.ReadAllLines(filePath);
                                added = 0;
                                skipped = 0;

                              


                                // Keep track of IDs being added in this session to avoid duplicates within the bulk file
                                currentBulkIDs = new HashSet<string>();

                                Console.WriteLine();

                                for (int i = 0; i < lines.Length; i++)
                                {
                                    line = lines[i].Trim();
                                    parts = line.Split(',');

                                    if (string.IsNullOrWhiteSpace(line))
                                        continue; // skip blank lines

                                    if (parts.Length == 6)
                                    {
                                        id = parts[0].Trim();
                                        lastName = parts[1].Trim();
                                        firstName = parts[2].Trim();

                                        // Parse grades directly as numbers
                                        g1 = parts[3].Trim();
                                        g2 = parts[4].Trim();
                                        g3 = parts[5].Trim();

                                        // Validate ID format
                                        idValid = false;

                                        if (id.Length == 7 && id[0] == 'I' && id[1] == 'T' && id[2] == '1' &&
                                            char.IsLetter(id[3]) && char.IsDigit(id[4]) && char.IsDigit(id[5]) && char.IsDigit(id[6]))
                                        {
                                            idValid = true;
                                        }

                                    


                                        if (!idValid)
                                        {
                                            skipped++;
                                            Console.ForegroundColor = ConsoleColor.Red;
                                            Console.WriteLine($"Line {i + 1} skipped: Invalid student ID format (expected: IT1A005).");
                                            Console.ForegroundColor = ConsoleColor.White;
                                            continue;
                                        }

                                        // Validate names (letters only, proper capitalization)
                                        nameValid = true;
                                        if (string.IsNullOrWhiteSpace(lastName) || lastName.Any(char.IsDigit) ||
                                            !char.IsUpper(lastName[0]) || lastName.Substring(1) != lastName.Substring(1).ToLower() ||
                                            string.IsNullOrWhiteSpace(firstName) || firstName.Any(char.IsDigit) ||
                                            !char.IsUpper(firstName[0]) || firstName.Substring(1) != firstName.Substring(1).ToLower())
                                        {
                                            nameValid = false;
                                        }

                                        if (!nameValid)
                                        {
                                            skipped++;
                                            Console.ForegroundColor = ConsoleColor.Red;
                                            Console.WriteLine($"Line {i + 1} skipped: Invalid name format (names must start with capital letter, contain only letters).");
                                            Console.ForegroundColor = ConsoleColor.White;
                                            continue;
                                        }

                                        // Check for duplicate ID in existing records
                                        duplicate = false;
                                        foreach (string existingRecord in existingRecords)
                                        {
                                            if (existingRecord.StartsWith(id + ","))
                                            {
                                                duplicate = true;
                                                break;
                                            }
                                        }

                                        // Check for duplicate ID in current bulk operation
                                        if (currentBulkIDs.Contains(id))
                                        {
                                            duplicate = true;
                                        }

                                        if (duplicate)
                                        {
                                            skipped++;
                                            Console.ForegroundColor = ConsoleColor.Red;
                                            Console.WriteLine($"Line {i + 1} skipped: Duplicate student ID ({id}).");
                                            Console.ForegroundColor = ConsoleColor.White;
                                            continue;
                                        }

                                        // Parse and validate grades
                                        g1Valid = double.TryParse(g1, out grade1);
                                        g2Valid = double.TryParse(g2, out grade2);
                                        g3Valid = double.TryParse(g3, out grade3);

                                        if (!g1Valid || !g2Valid || !g3Valid)
                                        {
                                            skipped++;
                                            Console.ForegroundColor = ConsoleColor.Red;
                                            Console.WriteLine($"Line {i + 1} skipped: Invalid grade format (grades must be numbers).");
                                            Console.ForegroundColor = ConsoleColor.White;
                                            continue;
                                        }

                                        // Validate grade range (70-100)
                                        if (grade1 < 70 || grade1 > 100 || grade2 < 70 || grade2 > 100 || grade3 < 70 || grade3 > 100)
                                        {
                                            skipped++;
                                            Console.ForegroundColor = ConsoleColor.Red;
                                            Console.WriteLine($"Line {i + 1} skipped: Grades must be between 70-100.");
                                            Console.ForegroundColor = ConsoleColor.White;
                                            continue;
                                        }

                                        // All validations passed - add the record
                                        record = $"{id},{lastName},{firstName},{grade1},{grade2},{grade3}";

                                        try
                                        {
                                            // Ensure directory exists before writing
                                           directory = Path.GetDirectoryName(dataPath);
                                            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                                            {
                                                Directory.CreateDirectory(directory);
                                            }

                                            File.AppendAllText(dataPath, record + "\n"); // Append the record to the CSV file

                                            // Add to tracking collections
                                            studentIDs.Add(id);
                                            studentLastNames.Add(lastName);
                                            studentFirstNames.Add(firstName);
                                            currentBulkIDs.Add(id);

                                            added++;
                                            Console.ForegroundColor = ConsoleColor.Green;
                                            Console.WriteLine($"Line {i + 1} added: {lastName}, {firstName} ({id})");
                                            Console.ForegroundColor = ConsoleColor.White;
                                        }
                                        catch (UnauthorizedAccessException)
                                        {
                                            Console.ForegroundColor = ConsoleColor.Red;
                                            Console.WriteLine($"Line {i + 1} skipped: Access denied while saving record.");
                                            Console.ForegroundColor = ConsoleColor.White;
                                            skipped++;
                                        }
                                        catch (IOException ex)
                                        {
                                            Console.ForegroundColor = ConsoleColor.Red;
                                            Console.WriteLine($"Line {i + 1} skipped: Error saving record - {ex.Message}");
                                            Console.ForegroundColor = ConsoleColor.White;
                                            skipped++;
                                        }
                                    }
                                    else
                                    {
                                        skipped++;
                                        Console.ForegroundColor = ConsoleColor.Red;
                                        Console.WriteLine($"Line {i + 1} skipped: Incorrect format (expecting 6 fields: ID,LastName,FirstName,Grade1,Grade2,Grade3).");
                                        Console.ForegroundColor = ConsoleColor.White;
                                    }
                                }

                                // Sort the records after adding all new ones
                                if (added > 0)
                                {
                                    try
                                    {
                                        allRecords = File.ReadAllLines(dataPath);
                                        Array.Sort(allRecords);
                                        File.WriteAllLines(dataPath, allRecords);
                                    }
                                    catch (IOException ex)
                                    {
                                        Console.ForegroundColor = ConsoleColor.Yellow;
                                        Console.WriteLine($"Warning: Unable to sort records - {ex.Message}");
                                        Console.ForegroundColor = ConsoleColor.White;
                                    }
                                }

                                Console.ForegroundColor = ConsoleColor.Green;
                                Console.WriteLine($"\nBulk import completed!");
                                Console.ForegroundColor = ConsoleColor.DarkGreen;
                                Console.WriteLine($"\nSuccessfully added: {added} records");
                                Console.ForegroundColor = ConsoleColor.DarkRed;
                                Console.WriteLine($"Skipped: {skipped} records");
                                Console.ForegroundColor = ConsoleColor.White;
                            }
                            catch (IOException ex)
                            {
                                Console.ForegroundColor = ConsoleColor.Red;
                                Console.WriteLine($"Error reading bulk import file: {ex.Message}");
                                Console.ForegroundColor = ConsoleColor.White;
                            }
                            catch (UnauthorizedAccessException)
                            {
                                Console.ForegroundColor = ConsoleColor.Red;
                                Console.WriteLine("Access denied: Unable to read the bulk import file. Please check file permissions.");
                                Console.ForegroundColor = ConsoleColor.White;
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine($"Unexpected error during bulk import: {ex.Message}");
                            Console.ForegroundColor = ConsoleColor.White;
                        }
                        break;

                    case "3": // Export all students
                        try
                        {
                            Console.WriteLine(); // spacing

                            // Ensure Report folder exists
                            try
                            {
                                if (!Directory.Exists("Report"))
                                    Directory.CreateDirectory("Report");
                            }
                            catch (UnauthorizedAccessException)
                            {
                                Console.ForegroundColor = ConsoleColor.Red;
                                Console.WriteLine("Access denied: Unable to create Report directory. Please check permissions.");
                                Console.ForegroundColor = ConsoleColor.White;
                                break;
                            }
                            catch (IOException ex)
                            {
                                Console.ForegroundColor = ConsoleColor.Red;
                                Console.WriteLine($"Error creating Report directory: {ex.Message}");
                                Console.ForegroundColor = ConsoleColor.White;
                                break;
                            }

                            // Read all records from the data file
                            try
                            {
                                records = File.ReadAllLines(dataPath);
                            }
                            catch (FileNotFoundException)
                            {
                                Console.ForegroundColor = ConsoleColor.Red;
                                Console.WriteLine("No student data found. Please add students first.");
                                Console.ForegroundColor = ConsoleColor.White;
                                break;
                            }
                            catch (IOException ex)
                            {
                                Console.ForegroundColor = ConsoleColor.Red;
                                Console.WriteLine($"Error reading student data: {ex.Message}");
                                Console.ForegroundColor = ConsoleColor.White;
                                break;
                            }

                            // Create date-stamped filename
                            date = DateTime.Now.ToString("yyyyMMdd");
                            reportFile = reportsPath + date + ".csv";

                            try
                            {
                                // Delete if it already exists
                                if (File.Exists(reportFile))
                                {
                                    Console.ForegroundColor = ConsoleColor.Red;
                                    Console.WriteLine("Report file already exists. Deleting old file...");
                                    Thread.Sleep(1500);
                                    File.Delete(reportFile);
                                }
                                else // First time making the file
                                {
                                    Thread.Sleep(1000);
                                    Console.WriteLine("Creating new report file...");
                                    Thread.Sleep(1500);
                                }

                                

                                // Append all student records

                                // Write CSV header
                                File.WriteAllText(reportFile, "Student ID,Last Name,First Name,Data Structures,Programming 2,Math Applications in IT,Average Grade\n");

                                // This loop just adds average score column ( does not reflect on case 1)

                                // Use a HashSet to track if a header was already written
                               sectionsWritten = new HashSet<string>();

                                for (int i = 0; i < records.Length; i++)
                                {
                                    parts = records[i].Split(',');

                                    if (parts.Length >= 6)
                                    {
                                        double Grade1, Grade2, Grade3, Avg;
                                        bool valid1 = double.TryParse(parts[3], out Grade1);
                                        bool valid2 = double.TryParse(parts[4], out Grade2);
                                        bool valid3 = double.TryParse(parts[5], out Grade3);

                                        if (valid1 && valid2 && valid3)
                                        {
                                            Avg = (Grade1 + Grade2 + Grade3) / 3;

                                            section = parts[0].Substring(3, 1).ToUpper(); // Gets 'A', 'B', or 'C' from ID like IT1A001

                                            // Add section header if not already written
                                            if (!sectionsWritten.Contains(section))
                                            {
                                                File.AppendAllText(reportFile, $"\nBSIT 1{section},---------------------,---------------------,---------------------,---------------------,---------------------,---------------------\n");
                                                sectionsWritten.Add(section);
                                            }

                                            updatedLine = $"{parts[0]},{parts[1]},{parts[2]},{Grade1},{Grade2},{Grade3},{Avg:F2}";
                                            File.AppendAllText(reportFile, updatedLine + "\n");
                                        }
                                    }
                                }
                            


                                // Confirm export
                                Console.ForegroundColor = ConsoleColor.Green;
                                Console.WriteLine("\nExport successful. File saved to: " + Path.GetFullPath(reportFile) + "\n");
                                Console.ForegroundColor = ConsoleColor.White;
                            }
                            catch (UnauthorizedAccessException)
                            {
                                Console.ForegroundColor = ConsoleColor.Red;
                                Console.WriteLine("Access denied: Unable to create or write report file. Please check permissions.");
                                Console.ForegroundColor = ConsoleColor.White;
                            }
                            catch (PathTooLongException)
                            {
                                Console.ForegroundColor = ConsoleColor.Red;
                                Console.WriteLine("File path is too long. Please use a shorter path for the report.");
                                Console.ForegroundColor = ConsoleColor.White;
                            }
                            catch (IOException ex)
                            {
                                Console.ForegroundColor = ConsoleColor.Red;
                                Console.WriteLine($"Error creating report file: {ex.Message}");
                                Console.ForegroundColor = ConsoleColor.White;
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine($"Unexpected error during export: {ex.Message}");
                            Console.ForegroundColor = ConsoleColor.White;
                        }
                        break;

                    case "4": // Export individual student grades
                        try
                        {
                            Console.Write("\nEnter Student ID: ");
                            searchID = Console.ReadLine().Trim();
                            Console.WriteLine(); // spacing

                            // Read student data
                            try
                            {
                                allLines = File.ReadAllLines(dataPath);
                                found = false;

                                foreach (string searchLine in allLines)
                                {
                                    p = searchLine.Split(',');

                                    if (p.Length == 6 && p[0] == searchID) // Check if the line has 6 parts and matches the search ID
                                    {
                                        found = true;

                                        lastName = p[1].Trim(); // Access last name and trim whitespace
                                        firstName = p[2].Trim(); // Access first name and trim whitespace

                                        id = p[0].Trim(); // Access student ID and trim whitespace

                                        try
                                        {
                                            grade1 = double.Parse(p[3]); // Access grades and parse to double
                                            grade2 = double.Parse(p[4]); // Access grades and parse to double
                                            grade3 = double.Parse(p[5]); // Access grades and parse to double

                                            avg = (grade1 + grade2 + grade3) / 3; // Calculate average grade

                                            // Ensure Report folder exists
                                            try
                                            {
                                                if (!Directory.Exists("Report"))
                                                {
                                                    Directory.CreateDirectory("Report");
                                                }
                                            }
                                            catch (UnauthorizedAccessException)
                                            {
                                                Console.ForegroundColor = ConsoleColor.Red;
                                                Console.WriteLine("Access denied: Unable to create Report directory. Please check permissions.");
                                                Console.ForegroundColor = ConsoleColor.White;
                                                break;
                                            }
                                            catch (IOException ex)
                                            {
                                                Console.ForegroundColor = ConsoleColor.Red;
                                                Console.WriteLine($"Error creating Report directory: {ex.Message}");
                                                Console.ForegroundColor = ConsoleColor.White;
                                                break;
                                            }

                                            studentFile = $"Report/{lastName}, {firstName}_Grades.csv";

                                            try
                                            {
                                                if (File.Exists(studentFile))
                                                {
                                                    File.Delete(studentFile);
                                                }

                                                // Write student grade report
                                                File.AppendAllText(studentFile, $"ID,Name,Data Structures,Programming 2,Math Application in IT,Average Grade\n");
                                                File.AppendAllText(studentFile,
                                                    $"{id},{firstName} {lastName},{grade1},{grade2},{grade3},{avg:F2}\n");


                                                Console.ForegroundColor = ConsoleColor.Green;
                                                Console.WriteLine("Export successful. File saved to: " + Path.GetFullPath(studentFile) + "\n");
                                                Console.ForegroundColor = ConsoleColor.White;
                                            }
                                            catch (UnauthorizedAccessException)
                                            {
                                                Console.ForegroundColor = ConsoleColor.Red;
                                                Console.WriteLine("Access denied: Unable to create individual student report. Please check permissions.");
                                                Console.ForegroundColor = ConsoleColor.White;
                                            }
                                            catch (PathTooLongException)
                                            {
                                                Console.ForegroundColor = ConsoleColor.Red;
                                                Console.WriteLine("File path is too long. Please use shorter names or paths.");
                                                Console.ForegroundColor = ConsoleColor.White;
                                            }
                                            catch (IOException ex)
                                            {
                                                Console.ForegroundColor = ConsoleColor.Red;
                                                Console.WriteLine($"Error creating individual student report: {ex.Message}");
                                                Console.ForegroundColor = ConsoleColor.White;
                                            }
                                        }
                                        catch (FormatException)
                                        {
                                            Console.ForegroundColor = ConsoleColor.Red;
                                            Console.WriteLine("Error: Invalid grade format in student record. Data may be corrupted.");
                                            Console.ForegroundColor = ConsoleColor.White;
                                        }
                                        break;
                                    }
                                }

                                if (!found)
                                {
                                    Console.ForegroundColor = ConsoleColor.Red;
                                    Console.WriteLine("Student not found.\n");
                                    Console.ForegroundColor = ConsoleColor.White;
                                }
                            }
                            catch (FileNotFoundException)
                            {
                                Console.ForegroundColor = ConsoleColor.Red;
                                Console.WriteLine("Student data file not found. Please add students first.");
                                Console.ForegroundColor = ConsoleColor.White;
                            }
                            catch (IOException ex)
                            {
                                Console.ForegroundColor = ConsoleColor.Red;
                                Console.WriteLine($"Error reading student data: {ex.Message}");
                                Console.ForegroundColor = ConsoleColor.White;
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine($"Unexpected error during individual export: {ex.Message}");
                            Console.ForegroundColor = ConsoleColor.White;
                        }
                        break;

                    case "5": // Exit program
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        endProgram = true; // Set the flag to true to exit the loop
                        Console.WriteLine("\nExiting the program.\n");
                        break;

                    default: // Invalid option handling
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("\nInvalid option. Please select a valid menu option (1-5).\n");
                        Console.ForegroundColor = ConsoleColor.White;
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"\nUnexpected error in main menu: {ex.Message}");
                Console.ForegroundColor = ConsoleColor.White;
            }

            // Return to menu unless exiting
            if (!endProgram)
            {
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("\nPress Enter to return to the menu...");
                Console.ReadLine();
                Console.Clear();
            }
        }

        // Program cleanup before exit
        try
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("Thank you for using the Student Grade Manager!");
            Console.ForegroundColor = ConsoleColor.White;
        }
        catch (Exception)
        {
            // Silent catch for any final cleanup issues
        }
        Console.WriteLine();
    }
}
