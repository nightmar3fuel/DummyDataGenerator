using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

class Program
{
    static Random random = new Random();

    
    static string[] bachelorChoices = { "CS", "AMC", "IS" };

    // Departments for teachers
    static string[] departments = { "College of Business", "College of Arts and Sciences", "College of Computer and Information Science" };

    // Subjects for teachers
    static string[] subjects = { "Mathematics", "Computer Science", "Marketing", "English", "Physics", "Economics", "Programming", "Database Management" };

    
    static List<string> lastNames = new List<string>
    {
        "Dela Cruz", "García", "Reyes", "Ramos", "Mendoza", "Santos", "Flores", "Gonzales", "Bautista", "Villanueva",
        "Najera", "Cruz", "de Guzmán", "López", "Pérez", "Castillo", "Francisco", "Rivera", "Aquino", "Castro",
        "Sánchez", "Torres", "de León", "Domingo", "Martínez", "Rodríguez", "Santiago", "Soriano", "delos Santos",
        "Díaz", "Hernández", "Tolentino", "Valdez", "Ramírez", "Morales", "Mercado", "Tan", "Aguilar", "Navarro",
        "Manalo", "Gómez", "Dizon", "del Rosario", "Javier", "Córpuz", "Gutiérrez", "Salvador", "Velasco", "Miranda",
        "David"
    };

    
    static List<string> firstNames = new List<string>
    {
        // Male Names
        "Juan", "Jose", "Miguel", "Rafael", "David", "Daniel", "Elias", "Pedro", "Tomas", "Lorenzo",
        "James", "Michael", "Robert", "John", "William", "Richard", "Joseph", "Thomas", "Christopher", 
        
        // Female Names
        "Maria", "Ana", "Isabel", "Carla", "Elena", "Sofia", "Carmela", "Mary", "Patricia", "Jennifer",
        "Linda", "Elizabeth", "Barbara", "Susan", "Jessica", "Karen", "Sarah"
    };

    
    static List<char> middleInitials = Enumerable.Range('A', 26).Select(i => (char)i).ToList();

    
    static HashSet<string> emailSet = new HashSet<string>();

    static void Main()
    {
        List<string[]> studentData = new List<string[]>();
        List<string[]> teacherData = new List<string[]>();

        // Generate student data
        for (int i = 0; i < 100; i++)
        {
            studentData.Add(GeneratePersonData(bachelorChoices));
        }

        // Generate teacher data
        for (int i = 0; i < 50; i++)
        {
            teacherData.Add(GenerateTeacherData());
        }

        // Export data to Excel file
        ExportToExcel(studentData, teacherData);
    }

    // Function to generate unique 11-digit code
    static string GenerateUniqueCode()
    {
        return string.Join("", Enumerable.Range(0, 11).Select(_ => random.Next(10).ToString()));
    }

    // Function to generate student or teacher data
    static string[] GeneratePersonData(string[] bachelorsOrSubjects)
    {
        string lastName = lastNames[random.Next(lastNames.Count)];
        string firstName = firstNames[random.Next(firstNames.Count)];
        char middleInitial = middleInitials[random.Next(middleInitials.Count)];
        string choice = bachelorsOrSubjects[random.Next(bachelorsOrSubjects.Length)];
        string uniqueCode = GenerateUniqueCode();

        // Generate email
        string emailBase = $"{firstName[0].ToString().ToLower()}{lastName.Replace(" ", "").ToLower()}";
        string email = $"{emailBase}@mcm.edu.ph";

        // Ensure email uniqueness
        int count = 1;
        while (emailSet.Contains(email))
        {
            count++;
            email = $"{firstName.Substring(0, count).ToLower()}{lastName.Replace(" ", "").ToLower()}@mcm.edu.ph";
        }

        emailSet.Add(email);

        return new string[] { lastName, firstName, middleInitial.ToString(), uniqueCode, choice, email };
    }

    // Function to generate teacher data
    static string[] GenerateTeacherData()
    {
        string lastName = lastNames[random.Next(lastNames.Count)];
        string firstName = firstNames[random.Next(firstNames.Count)];
        char middleInitial = middleInitials[random.Next(middleInitials.Count)];
        string subject = subjects[random.Next(subjects.Length)];
        string department = departments[random.Next(departments.Length)];
        string uniqueCode = GenerateUniqueCode();

        // Generate email
        string emailBase = $"{firstName[0].ToString().ToLower()}{lastName.Replace(" ", "").ToLower()}";
        string email = $"{emailBase}@mcm.edu.ph";

        // Ensure email uniqueness
        int count = 1;
        while (emailSet.Contains(email))
        {
            count++;
            email = $"{firstName.Substring(0, count).ToLower()}{lastName.Replace(" ", "").ToLower()}@mcm.edu.ph";
        }

        emailSet.Add(email);

        return new string[] { lastName, firstName, middleInitial.ToString(), uniqueCode, subject, department, email };
    }

    // Function to export data to an Excel file
    static void ExportToExcel(List<string[]> studentData, List<string[]> teacherData)
    {
        using (var workbook = new XLWorkbook())
        {
            var studentWorksheet = workbook.Worksheets.Add("Student Data");

            // Create headers for students
            studentWorksheet.Cell(1, 1).Value = "Last Name";
            studentWorksheet.Cell(1, 2).Value = "First Name";
            studentWorksheet.Cell(1, 3).Value = "Middle Initial";
            studentWorksheet.Cell(1, 4).Value = "Unique Code";
            studentWorksheet.Cell(1, 5).Value = "Bachelor";
            studentWorksheet.Cell(1, 6).Value = "Email";

            // Populate student data
            for (int i = 0; i < studentData.Count; i++)
            {
                studentWorksheet.Cell(i + 2, 1).Value = studentData[i][0];
                studentWorksheet.Cell(i + 2, 2).Value = studentData[i][1];
                studentWorksheet.Cell(i + 2, 3).Value = studentData[i][2];
                studentWorksheet.Cell(i + 2, 4).Value = studentData[i][3];
                studentWorksheet.Cell(i + 2, 5).Value = studentData[i][4];
                studentWorksheet.Cell(i + 2, 6).Value = studentData[i][5];
            }

            // Adjust column widths for readability
            studentWorksheet.Columns().AdjustToContents();

            var teacherWorksheet = workbook.Worksheets.Add("Teacher Data");

            // Create headers for teachers
            teacherWorksheet.Cell(1, 1).Value = "Last Name";
            teacherWorksheet.Cell(1, 2).Value = "First Name";
            teacherWorksheet.Cell(1, 3).Value = "Middle Initial";
            teacherWorksheet.Cell(1, 4).Value = "Unique Code";
            teacherWorksheet.Cell(1, 5).Value = "Subject";
            teacherWorksheet.Cell(1, 6).Value = "Department";
            teacherWorksheet.Cell(1, 7).Value = "Email";

            // Populate teacher data
            for (int i = 0; i < teacherData.Count; i++)
            {
                teacherWorksheet.Cell(i + 2, 1).Value = teacherData[i][0];
                teacherWorksheet.Cell(i + 2, 2).Value = teacherData[i][1];
                teacherWorksheet.Cell(i + 2, 3).Value = teacherData[i][2];
                teacherWorksheet.Cell(i + 2, 4).Value = teacherData[i][3];
                teacherWorksheet.Cell(i + 2, 5).Value = teacherData[i][4];
                teacherWorksheet.Cell(i + 2, 6).Value = teacherData[i][5];
                teacherWorksheet.Cell(i + 2, 7).Value = teacherData[i][6];
            }

            // Adjust column widths for readability
            teacherWorksheet.Columns().AdjustToContents();

            // Save the Excel file
            workbook.SaveAs("DummyData.xlsx");
            Console.WriteLine("Excel file saved as 'DummyData.xlsx'");
        }
    }
}
