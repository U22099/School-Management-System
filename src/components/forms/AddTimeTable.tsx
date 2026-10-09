import React, { useState, useMemo } from "react";
import { motion } from "framer-motion";
import * as XLSX from "xlsx";
import { ArrowLeft, ArrowRight, Download, LoaderCircle } from "lucide-react";
import toast from "react-hot-toast";

const AddTimeTableForm = ({
  setTotalData,
  setPage,
  defaultValues,
  submit,
  loading,
}: {
  setTotalData: React.Dispatch<React.SetStateAction<TotalData>>;
  setPage: React.Dispatch<React.SetStateAction<number>>;
  defaultValues: { timetableData: Timetable[] | []; timetableHtml: string };
  submit: () => Promise<void>;
  loading: boolean;
}) => {
  const [timetableData, setTimetableData] = useState<Timetable[]>(
    defaultValues.timetableData
  );
  const [timetableHtml, setTimetableHtml] = useState<string>(
    defaultValues.timetableHtml
  );
  const [fileError, setFileError] = useState<string | null>(null);
  const [fileLoading, setFileLoading] = useState<boolean>(false);

  const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    setFileLoading(true);
    try {
      const file = event.target.files?.[0];
      if (!file) return;

      const validExtensions = [".xlsx", ".xls", ".xlsm"];
      const fileExtension = file.name.slice(file.name.lastIndexOf("."));
      if (!validExtensions.includes(fileExtension)) {
        toast("File is not a valid Excel");
        setFileError("Invalid file type. Please upload an Excel file.");
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
        const jsonData: any[] = XLSX.utils.sheet_to_json(sheet, {
          header: 1,
        });
        const htmlData: string = XLSX.utils.sheet_to_html(sheet);
        const extractedData: Timetable[] = [];
        let lastDay = "";
        const periods = jsonData[0]?.slice(2) || [];

        jsonData.forEach((row, rowIndex) => {
          if (rowIndex === 0) return;
          let day = lastDay;
          if (row[0]) {
            day = row[0];
            lastDay = row[0];
          }

          if (day) {
            let period = 1;
            const subjects = row.slice(2);

            for (let colIndex = 0; colIndex < subjects.length; colIndex++) {
              if (subjects[colIndex]) {
                extractedData.push({
                  class: row[1],
                  day,
                  period: period++,
                  periodSpan: 1,
                  subject: subjects[colIndex].toString(),
                  startTime: periods[colIndex]?.split(" - ")[0] || "",
                  endTime: periods[colIndex]?.split(" - ")[1] || "",
                });
              } else {
                extractedData[extractedData.length - 1].endTime =
                  periods[colIndex]?.split(" - ")[1] || "";
                extractedData[extractedData.length - 1].periodSpan += 1;
              }
            }
          }
        });
        if (extractedData.length < 1) {
          toast("Timetable is empty");
        } else {
          setTimetableData(extractedData);
          setTimetableHtml(htmlData);
          setTotalData((prev) => ({
            ...prev,
            timetable: extractedData,
            timetableHtml: htmlData,
          }));
        }
      };

      reader.readAsBinaryString(file);
    } catch (err: any) {
      console.error("Error occurred when reading file", err);
      setFileError(err?.message || "An unknown error occurred");
    } finally {
      setFileLoading(false);
    }
  };

  return (
    <motion.div
      initial={{ opacity: 0 }}
      animate={{ opacity: 1 }}
      transition={{ duration: 0.6 }}
      className="p-6 rounded-lg w-full h-full flex flex-col gap-4 mb-4 mt-5"
    >
      <h1 className="text-2xl font-bold mb-4">Add Timetable</h1>
      {!timetableHtml && (
        <div className="flex flex-col gap-2 w-full">
          <h2 className="text-xl font-bold">File Sample</h2>
          <FileFormatSample />
        </div>
      )}
      <div className="w-full items-center md:justify-between px-2 flex gap-2 flex-col md:flex-row">
        <button
          onClick={() => {
            const a = document.createElement("a");
            a.download = "Timetable's Excel Template";
            a.href = "/timetableExcelTemp.xlsx";
            a.click();
          }}
          disabled={loading}
          className="button w-full md:w-fit"
        >
          <Download className="stroke-black" /> Download Excel Template
        </button>
        <input
          type="file"
          accept=".xlsx,.xls,.xlsm"
          onChange={handleFileUpload}
          disabled={loading}
          className="input flex w-full md:w-fit text-sm text-primary p-2 bg-secondary rounded-lg file:mr-4 file:py-2 file:px-4 file:rounded file:border-0 file:text-sm file:font-semibold file:bg-secondary file:text-primary hover:file:bg-secondary file:cursor-pointer cursor-pointer"
        />
      </div>
      {fileLoading && (
        <p className="text-green-500 text-sm mb-4 animate-pulse">Loading...</p>
      )}
      {fileError && (
        <p className="text-red-500 text-sm mb-4 font-bold">{fileError}</p>
      )}
      {timetableHtml && (
        <div
          className="w-full overflow-x-scroll mx-auto"
          dangerouslySetInnerHTML={{ __html: timetableHtml }}
        />
      )}
      <div className="w-full flex gap-4 justify-between">
        <button
          disabled={loading}
          onClick={() => setPage((prev) => prev - 1)}
          className="button"
        >
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
                  (i === 6 ? "bg-secondary" : "bg-transparent")
                }
              />
            ))}
        </div>
        <button
          disabled={timetableData.length < 1 || loading}
          onClick={!(timetableData.length < 1 || loading) ? submit : () => console.log("Nice Try")}
          className="button justify-end"
        >
          {loading && <LoaderCircle className="animate-spin" />} Finish{" "}
          <ArrowRight className="stroke-black" />
        </button>
      </div>
    </motion.div>
  );
};

export default AddTimeTableForm;

const FileFormatSample = () => {
  return (
    <div
      className="w-full overflow-x-scroll mx-auto"
      dangerouslySetInnerHTML={{ __html: htmlSampleTable }}
    />
  );
};

/*  Sample HTML Table */
const htmlSampleTable: string = `<html><head><meta charset="utf-8"/><title>SheetJS Table Export</title></head><body><table><tr><td data-t="s" data-v="TIMETABLE" id="sjs-A1">TIMETABLE</td><td id="sjs-B1"></td><td data-t="s" data-v="8:10AM - 8:50AM" id="sjs-C1">8:10AM - 8:50AM</td><td data-t="s" data-v="8:50AM - 9:30AM" id="sjs-D1">8:50AM - 9:30AM</td><td data-t="s" data-v="9:30AM - 10:10AM" id="sjs-E1">9:30AM - 10:10AM</td><td data-t="s" data-v="10:10AM - 10:50AM" id="sjs-F1">10:10AM - 10:50AM</td><td data-t="s" data-v="10:50AM - 11:30AM" id="sjs-G1">10:50AM - 11:30AM</td><td data-t="s" data-v="12:00PM - 12:40PM" id="sjs-H1">12:00PM - 12:40PM</td><td data-t="s" data-v="12:40PM - 1:20PM" id="sjs-I1">12:40PM - 1:20PM</td><td data-t="s" data-v="1:20PM - 2:00PM" id="sjs-J1">1:20PM - 2:00PM</td></tr><tr><td rowspan="2" data-t="s" data-v="MONDAY" id="sjs-A2">MONDAY</td><td data-t="s" data-v="1A" id="sjs-B2">1A</td><td colspan="2" data-t="s" data-v="MATHEMATICS" id="sjs-C2">MATHEMATICS</td><td data-t="s" data-v="CHEMISTRY" id="sjs-E2">CHEMISTRY</td><td colspan="2" data-t="s" data-v="PHYSICS" id="sjs-F2">PHYSICS</td><td colspan="2" data-t="s" data-v="FURTHER MATHS" id="sjs-H2">FURTHER MATHS</td><td data-t="s" data-v="ANIMAL HUSBANDARY" id="sjs-J2">ANIMAL HUSBANDARY</td></tr><tr><td data-t="s" data-v="1B" id="sjs-B3">1B</td><td data-t="s" data-v="MARKETING" id="sjs-C3">MARKETING</td><td colspan="2" data-t="s" data-v="COMPUTER" id="sjs-D3">COMPUTER</td><td colspan="2" data-t="s" data-v="ACCOUNTING" id="sjs-F3">ACCOUNTING</td><td colspan="2" data-t="s" data-v="COMMERCE" id="sjs-H3">COMMERCE</td><td data-t="s" data-v="CIVIC" id="sjs-J3">CIVIC</td></tr><tr><td rowspan="2" data-t="s" data-v="TUESDAY" id="sjs-A4">TUESDAY</td><td data-t="s" data-v="1A" id="sjs-B4">1A</td><td colspan="2" data-t="s" data-v="MATHEMATICS" id="sjs-C4">MATHEMATICS</td><td data-t="s" data-v="CHEMISTRY" id="sjs-E4">CHEMISTRY</td><td colspan="2" data-t="s" data-v="PHYSICS" id="sjs-F4">PHYSICS</td><td colspan="2" data-t="s" data-v="FURTHER MATHS" id="sjs-H4">FURTHER MATHS</td><td data-t="s" data-v="ANIMAL HUSBANDARY" id="sjs-J4">ANIMAL HUSBANDARY</td></tr><tr><td data-t="s" data-v="1B" id="sjs-B5">1B</td><td data-t="s" data-v="MARKETING" id="sjs-C5">MARKETING</td><td colspan="2" data-t="s" data-v="COMPUTER" id="sjs-D5">COMPUTER</td><td colspan="2" data-t="s" data-v="ACCOUNTING" id="sjs-F5">ACCOUNTING</td><td colspan="2" data-t="s" data-v="COMMERCE" id="sjs-H5">COMMERCE</td><td data-t="s" data-v="CIVIC" id="sjs-J5">CIVIC</td></tr><tr><td rowspan="2" data-t="s" data-v="WEDNESDAY" id="sjs-A6">WEDNESDAY</td><td data-t="s" data-v="1A" id="sjs-B6">1A</td><td colspan="2" data-t="s" data-v="MATHEMATICS" id="sjs-C6">MATHEMATICS</td><td data-t="s" data-v="CHEMISTRY" id="sjs-E6">CHEMISTRY</td><td colspan="2" data-t="s" data-v="PHYSICS" id="sjs-F6">PHYSICS</td><td colspan="2" data-t="s" data-v="FURTHER MATHS" id="sjs-H6">FURTHER MATHS</td><td data-t="s" data-v="ANIMAL HUSBANDARY" id="sjs-J6">ANIMAL HUSBANDARY</td></tr><tr><td data-t="s" data-v="1B" id="sjs-B7">1B</td><td data-t="s" data-v="MARKETING" id="sjs-C7">MARKETING</td><td colspan="2" data-t="s" data-v="COMPUTER" id="sjs-D7">COMPUTER</td><td colspan="2" data-t="s" data-v="ACCOUNTING" id="sjs-F7">ACCOUNTING</td><td colspan="2" data-t="s" data-v="COMMERCE" id="sjs-H7">COMMERCE</td><td data-t="s" data-v="CIVIC" id="sjs-J7">CIVIC</td></tr><tr><td rowspan="2" data-t="s" data-v="THURSDAY" id="sjs-A8">THURSDAY</td><td data-t="s" data-v="1A" id="sjs-B8">1A</td><td colspan="2" data-t="s" data-v="MATHEMATICS" id="sjs-C8">MATHEMATICS</td><td data-t="s" data-v="CHEMISTRY" id="sjs-E8">CHEMISTRY</td><td colspan="2" data-t="s" data-v="PHYSICS" id="sjs-F8">PHYSICS</td><td colspan="2" data-t="s" data-v="FURTHER MATHS" id="sjs-H8">FURTHER MATHS</td><td data-t="s" data-v="ANIMAL HUSBANDARY" id="sjs-J8">ANIMAL HUSBANDARY</td></tr><tr><td data-t="s" data-v="1B" id="sjs-B9">1B</td><td data-t="s" data-v="MARKETING" id="sjs-C9">MARKETING</td><td colspan="2" data-t="s" data-v="COMPUTER" id="sjs-D9">COMPUTER</td><td colspan="2" data-t="s" data-v="ACCOUNTING" id="sjs-F9">ACCOUNTING</td><td colspan="2" data-t="s" data-v="COMMERCE" id="sjs-H9">COMMERCE</td><td data-t="s" data-v="CIVIC" id="sjs-J9">CIVIC</td></tr><tr><td rowspan="2" data-t="s" data-v="FRIDAY" id="sjs-A10">FRIDAY</td><td data-t="s" data-v="1A" id="sjs-B10">1A</td><td colspan="2" data-t="s" data-v="MATHEMATICS" id="sjs-C10">MATHEMATICS</td><td data-t="s" data-v="CHEMISTRY" id="sjs-E10">CHEMISTRY</td><td colspan="2" data-t="s" data-v="PHYSICS" id="sjs-F10">PHYSICS</td><td colspan="2" data-t="s" data-v="FURTHER MATHS" id="sjs-H10">FURTHER MATHS</td><td data-t="s" data-v="ANIMAL HUSBANDARY" id="sjs-J10">ANIMAL HUSBANDARY</td></tr><tr><td data-t="s" data-v="1B" id="sjs-B11">1B</td><td data-t="s" data-v="MARKETING" id="sjs-C11">MARKETING</td><td colspan="2" data-t="s" data-v="COMPUTER" id="sjs-D11">COMPUTER</td><td colspan="2" data-t="s" data-v="ACCOUNTING" id="sjs-F11">ACCOUNTING</td><td colspan="2" data-t="s" data-v="COMMERCE" id="sjs-H11">COMMERCE</td><td data-t="s" data-v="CIVIC" id="sjs-J11">CIVIC</td></tr></table></body></html>
`;
