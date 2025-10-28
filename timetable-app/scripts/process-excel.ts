import XLSX from "xlsx";
import fs from "fs";
import path from "path";

interface TimetableData {
  course: string;
  color: string;
}

interface ProcessedSheet {
  [className: string]: TimetableData[][];
}

interface ProcessedData {
  [sheetName: string]: ProcessedSheet;
}

const DAYS_OF_WEEK = [
  "Timings",
  "Monday",
  "Tuesday",
  "Wednesday",
  "Thursday",
  "Friday",
];

const TIME_SLOTS = [
  "8:00am",
  "8:50am",
  "9:40am",
  "10:30am",
  "11:20am",
  "12:10pm",
  "1:00pm",
  "1:50pm",
  "2:40pm",
  "3:30pm",
  "4:20pm",
  "5:10pm",
  "6:00pm",
  "6:50pm",
];

function processExcelFile(filePath: string): ProcessedData {
  const workbook = XLSX.readFile(filePath);
  const processedData: ProcessedData = {};

  workbook.SheetNames.forEach((sheetName) => {
    const worksheet = workbook.Sheets[sheetName];
    const rawData = XLSX.utils.sheet_to_json(worksheet, {
      header: 1,
    }) as string[][];

    processedData[sheetName] = processSheetData(sheetName, rawData);
  });

  return processedData;
}

function processSheetData(sheetName: string, data: string[][]): ProcessedSheet {
  const classes = extractClasses(data);
  const processedSheet: ProcessedSheet = {};

  classes.forEach((className, classIndex) => {
    processedSheet[className] = extractTimetableForClass(data, classIndex);
  });

  return processedSheet;
}

function extractClasses(data: string[][]): string[] {
  const classRow = data[3] || [];
  return classRow
    .filter(
      (cell) =>
        cell &&
        cell !== "DAY" &&
        cell !== "HOURS" &&
        cell !== "SR NO" &&
        cell !== "SR.NO" &&
        cell !== "TUTORIAL"
    )
    .map((cell) => cell.trim());
}

function extractTimetableForClass(
  data: string[][],
  classIndex: number
): TimetableData[][] {
  const timetable: TimetableData[][] = [];

  const headerRow: TimetableData[] = DAYS_OF_WEEK.map((day) => ({
    course: day,
    color: "dark",
  }));
  timetable.push(headerRow);

  for (let i = 6; i < 147; i += 2) {
    const timeSlot = TIME_SLOTS[Math.floor((i - 6) / 2)] || "";
    const timeRow: TimetableData[] = [
      {
        course: timeSlot,
        color: "dark",
      },
    ];

    for (let dayIndex = 0; dayIndex < 5; dayIndex++) {
      const cellData = extractCellData(data, i, classIndex + 4, dayIndex);
      timeRow.push(cellData);
    }

    timetable.push(timeRow);
  }

  return timetable;
}

function extractCellData(
  data: string[][],
  row: number,
  col: number,
  dayOffset: number
): TimetableData {
  const cellValue = data[row]?.[col + dayOffset] || "";

  if (!cellValue.trim()) {
    return { course: "", color: "success" };
  }

  const processedCourse = processCourseCode(cellValue);
  const color = determineColor(cellValue);

  return {
    course: processedCourse,
    color,
  };
}

function processCourseCode(courseCode: string): string {
  let cleaned = courseCode.replace(/[\/\s]/g, "");

  if (cleaned.length > 6) {
    cleaned = cleaned.replace(/[LPT]$/, "");
  }

  const courseMapping: { [key: string]: string } = {};

  return courseMapping[cleaned] || courseCode;
}

function determineColor(courseCode: string): string {
  if (/^[A-Z]{3}[0-9]{3}\s?L/.test(courseCode)) {
    return "danger";
  } else if (/^[A-Z]{3}[0-9]{3}\s?T/.test(courseCode)) {
    return "primary";
  } else if (/^([A-Z]{3}[0-9]{3}(\/[A-Z]{3}[0-9]{3})+)\s?L/.test(courseCode)) {
    return "info";
  }
  return "success";
}

async function generateData() {
  try {
    const excelFilePath = "./timetable.xlsx";

    if (!fs.existsSync(excelFilePath)) {
      console.error("Excel file not found:", excelFilePath);
      console.log("Please add your timetable.xlsx file to the project root");
      return;
    }

    console.log("Processing Excel file...");
    const processedData = processExcelFile(excelFilePath);

    const dataDir = "./src/data";
    if (!fs.existsSync(dataDir)) {
      fs.mkdirSync(dataDir, { recursive: true });
    }

    const outputPath = path.join(dataDir, "timetable.json");
    fs.writeFileSync(outputPath, JSON.stringify(processedData, null, 2));

    console.log("Data processed successfully");
    console.log(
      `Generated data for ${Object.keys(processedData).length} sheets`
    );

    Object.entries(processedData).forEach(([sheetName, sheetData]) => {
      console.log(
        `   - ${sheetName}: ${Object.keys(sheetData).length} classes`
      );
    });
  } catch (error) {
    console.error("Error processing Excel file:", error);
    process.exit(1);
  }
}

generateData();
