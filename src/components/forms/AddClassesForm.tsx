import React, { useState } from "react";
import { motion } from "framer-motion";
import * as XLSX from "xlsx";
import { ArrowLeft, ArrowRight, Download } from "lucide-react";
import toast from "react-hot-toast";

const AddClassesForm = ({
  setTotalData,
  setPage,
  defaultValues,
}: {
  setTotalData: React.Dispatch<React.SetStateAction<TotalData>>;
  setPage: React.Dispatch<React.SetStateAction<number>>;
  defaultValues: ClassData[] | [];
}) => {
  const [classesData, setClassesData] = useState<ClassData[]>(defaultValues);
  const [fileError, setFileError] = useState<string | null>(null);
  const [fileLoading, setFileLoading] = useState<boolean>(false);

  const requiredColumns = ["name", "class teacher", "total student"];

  const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    setFileLoading(true);
    try {
      const file = event.target.files?.[0];
      if (!file) return;

      const validExtensions = [".xlsx", ".xls", ".xlsm"];
      const fileExtension = file.name.slice(file.name.lastIndexOf("."));
      if (!validExtensions.includes(fileExtension)) {
        toast("File is not a valid Excel");
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

        const parsedData: ClassData[] = jsonData.map((row) => ({
          name: row["NAME"] || "",
          classTeacher: row["CLASS TEACHER"] || "",
          totalStudent: row["TOTAL STUDENT"] || 0,
        }));

        const incompleteFields = parsedData
          .filter(
            (classData) =>
              !classData.name ||
              !classData.classTeacher ||
              !classData.totalStudent
          )
          .map((x) => x.name)
          .join();
        if (parsedData.length < 1) {
          toast("Class Data is empty");
        } else if (incompleteFields) {
          toast(`Incomplete fields for: ${incompleteFields}`);
        } else {
          setTotalData((prev) => ({ ...prev, classes: parsedData }));
          setClassesData(parsedData);
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
      className="p-6 w-full h-full flex flex-col justify-between gap-2 mb-4 mt-5"
    >
      <div className="w-full flex flex-col gap-4">
        <h1 className="text-2xl font-bold mb-4">Add Classes</h1>
        {classesData.length < 1 && (
          <div className="flex flex-col gap-2 w-full">
            <h2 className="text-xl font-bold">File Sample</h2>
            <FileFormatSample requiredColumns={requiredColumns} />
          </div>
        )}
        <div className="w-full items-center md:justify-between flex gap-2 flex-col md:flex-row">
          <button
            onClick={() => {
              const a = document.createElement("a");
              a.download = "Class's Excel Template";
              a.href = "/classExcelTemp.xlsx";
              a.click();
            }}
            className="button w-full md:w-fit"
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
        {classesData.length > 0 && (
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
                {classesData.map((c, index) => (
                  <tr key={index}>
                    {Object.keys(c).map((col, i) => (
                      <td key={i}>{c[col as keyof ClassData]}</td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
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
                  (i === 4 ? "bg-secondary" : "bg-transparent")
                }
              />
            ))}
        </div>
        <button
  disabled={classesData.length < 1}
  onClick={() => setPage((prev) => prev + 1)}
  className="button justify-end"
>
  Next <ArrowRight className="stroke-black" />
</button>
      </div>
    </motion.div>
  );
};

export default AddClassesForm;

const FileFormatSample = ({
  requiredColumns,
}: {
  requiredColumns: string[];
}) => {
  const mockClassData: ClassData[] = [
    {
      name: "3A",
      classTeacher: "Mrs.Camille",
      totalStudent: 80,
    },
    {
      name: "3B",
      classTeacher: "Mrs.Camille",
      totalStudent: 80,
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
          {mockClassData.map((mockClass, index) => (
            <tr key={index}>
              {Object.keys(mockClass).map((col, i) => (
                <td key={i}>{mockClass[col as keyof ClassData]}</td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};
