import React, { useState } from "react";
import { motion } from "framer-motion";
import * as XLSX from "xlsx";
import { ArrowLeft, ArrowRight, Download } from "lucide-react";
import toast from "react-hot-toast";

const AddStudentsForm = ({
  setTotalData,
  setPage,
  defaultValues,
}: {
  setTotalData: React.Dispatch<React.SetStateAction<TotalData>>;
  setPage: React.Dispatch<React.SetStateAction<number>>;
  defaultValues: StudentData[] | [];
}) => {
  const [studentsData, setStudentsData] =
    useState<StudentData[]>(defaultValues);
  const [fileError, setFileError] = useState<string | null>(null);
  const [fileLoading, setFileLoading] = useState<boolean>(false);

  const requiredColumns = [
    "name",
    "email",
    "parent no",
    "parent name",
    "registration no",
    "doa",
    "birthdate",
    "admission no",
    "gender",
    "class",
    "address",
  ];

  const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    setFileLoading(true);
    try {
      const file = event.target.files?.[0];
      if (!file) return;

      const validExtensions = [".xlsx", ".xls", ".xlsm"];
      const fileExtension = file.name.slice(file.name.lastIndexOf("."));
      if (!validExtensions.includes(fileExtension)) {
        toast("File is not a valid Excel.");
        setFileError(
          "Invalid file type. Please upload an Excel file with the above format."
        );
        return;
      }

      setFileError(null);

      const reader = new FileReader();
      reader.onload = (e) => {
        const data = e.target?.result;
        if (!data) return;

        const workbook = XLSX.read(data, { type: "binary" });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const jsonData: any[] = XLSX.utils.sheet_to_json(sheet);

        const missingColumns = requiredColumns.filter(
          (col) => !Object.keys(jsonData[0] || {}).includes(col.toUpperCase())
        );

        if (missingColumns.length > 0) {
          toast("File Error");
          setFileError(
            `Missing required columns: ${missingColumns
              .join(", ")
              .toUpperCase()}. Please upload a valid Excel file.`
          );
          return;
        }

        const parsedData: StudentData[] = jsonData.map((row) => ({
          name: row["NAME"] || "",
          email: row["EMAIL"] || "",
          parentNo: row["PARENT NO"].toString() || "",
          parentName: row["PARENT NAME"] || "",
          registrationNo: row["REGISTRATION NO"].toString() || "",
          DOA: row["DOA"].toString() || "",
          birthdate: row["BIRTHDATE"].toString() || "",
          admissionNo: row["ADMISSION NO"].toString() || "",
          gender: row["GENDER"].toUpperCase() || "",
          class: row["CLASS"] || "",
          address: row["ADDRESS"] || "",
        }));

        const incompleteFields = parsedData
          .filter(
            (student) =>
              !student.name ||
              !student.email ||
              !student.parentName ||
              !student.parentNo
          )
          .map((x) => x.name || x.email)
          .join();
        if (parsedData.length < 1) {
          toast("Students Data is empty");
        } else if (incompleteFields) {
          toast(`Incomplete fields for: ${incompleteFields}`);
        } else {
          setTotalData((prev) => ({ ...prev, studentsData: parsedData }));
          setStudentsData(parsedData);
        }
      };

      reader.readAsBinaryString(file);
    } catch (err: any) {
      console.log("Error occured when reading file", err);
      setFileError(err?.message || "An unknown error occured");
    } finally {
      setFileLoading(false);
    }
  };

  return (
    <motion.div
      initial={{ opacity: 0 }}
      animate={{ opacity: 1 }}
      transition={{ duration: 0.6 }}
      className="p-6 rounded-lg w-full h-full flex flex-col justify-between gap-2 mb-4 mt-5"
    >
      <div className="w-full flex flex-col gap-4">
        <h1 className="text-2xl font-bold mb-4">Add Students</h1>
        {studentsData.length < 1 && (
          <div className="flex flex-col gap-2 w-full">
            <h2 className="text-xl font-bold">File Sample</h2>
            <FileFormatSample requiredColumns={requiredColumns} />
          </div>
        )}
        <div className="w-full items-center md:justify-between flex gap-2 flex-col md:flex-row">
          <button
            onClick={() => {
              const a = document.createElement("a");
              a.download = "Student's Excel Template";
              a.href = "/studentExcelTemp.xlsx";
              a.click();
            }}
            className="button w-full md:w-fit p-3"
          >
            <Download className="stroke-black" /> Download Excel Template
          </button>
          <input
            type="file"
            accept=".xlsx,.xls,.xlsm"
            onChange={handleFileUpload}
            className="input flex w-full md:w-fit text-sm text-primary p-2 bg-secondary rounded-lg file:mr-4 file:py-2 file:px-4 file:rounded file:border-0 file:text-sm file:font-semibold file:bg-secondary file:text-primary hover:file:bg-secondary file:cursor-pointer cursor-pointer"
          />
        </div>
        {fileLoading && (
          <p className="text-green-500 text-sm mb-4 animate-pulse">
            Loading...
          </p>
        )}
        {fileError && (
          <p className="text-red-500 text-sm mb-4 font-bold">{fileError}</p>
        )}
        {studentsData.length > 0 && (
          <div className="w-full overflow-x-scroll mx-auto">
            <table>
              <thead>
                <tr>
                  {requiredColumns.map((col, i) => (
                    <th key={i} className="capitalize">
                      {col}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {studentsData.map((student, index) => (
                  <tr key={index}>
                    {Object.keys(student).map((col, i) => (
                      <td key={i}>{student[col as keyof StudentData]}</td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
        <p className="text-xs font-semibold text-secondary-light">
          Make sure that the email section holds email of student and not
          parent!!
        </p>
      </div>
      <div className="w-full flex gap-4 justify-between">
        <button onClick={() => setPage((prev) => prev - 1)} className="button">
          <ArrowLeft className="stroke-black" /> Prev
        </button>
        <div className="flex gap-2 items-center">
          {Array(7)
            .fill("")
            .map((_, i) => (
              <div
                key={i}
                className={
                  "w-2 h-2 md:w-3 md:h-3 border border-secondary rounded-full " +
                  (i === 2 ? "bg-secondary" : "bg-transparent")
                }
              />
            ))}
        </div>
        <button
          disabled={studentsData.length < 1}
          onClick={() => {
  if (studentsData.length >= 1) {
    setPage((prev) => prev + 1);
  } else {
    console.log("Nice Try");
  }
}}
          className="button justify-end"
        >
          Next <ArrowRight className="stroke-black" />
        </button>
      </div>
    </motion.div>
  );
};

export default AddStudentsForm;

const FileFormatSample = ({
  requiredColumns,
}: {
  requiredColumns: string[];
}) => {
  const mockStudentData: StudentData[] = [
    {
      name: "John Doe",
      email: "john.doe@example.com",
      parentNo: "123-456-7890",
      parentName: "Jane Doe",
      registrationNo: "REG12345",
      DOA: "15-01-2022",
      birthdate: "10-11-2008",
      admissionNo: "ADM001",
      gender: "MALE",
      class: "3A",
      address: "123 Main Street, Cityville",
    },
    {
      name: "Alice Smith",
      email: "alice.smith@example.com",
      parentNo: "987-654-3210",
      parentName: "Bob Smith",
      registrationNo: "REG67890",
      DOA: "10-04-2023",
      birthdate: "15-09-2010",
      admissionNo: "ADM002",
      gender: "FEMALE",
      class: "2E",
      address: "456 Elm Street, Townsville",
    },
  ];
  return (
    <div className="w-full overflow-x-scroll mx-auto">
      <table>
        <thead>
          <tr>
            {requiredColumns.map((col, i) => (
              <th key={i} className="capitalize">
                {col}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {mockStudentData.map((student, index) => (
            <tr key={index}>
              {Object.keys(student).map((col, i) => (
                <td key={i}>{student[col as keyof StudentData]}</td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};
