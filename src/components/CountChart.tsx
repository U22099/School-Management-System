"use client";

import Image from "next/image";
import RadialCharts from "./RadialChart";
import { Mars, Venus } from "lucide-react";
import { useCounts } from "@/store";

const CountChart = () => {
  const counts = useCounts((state) => state.counts);
  const data = [
    {
      name: "Male",
      count: counts.maleStudents,
      fill: "#ceebfa",
    },
    {
      name: "Female",
      count: counts.femaleStudents,
      fill: "#fae27c",
    },
    {
      name: "Gender",
      count: counts.students,
      fill: "#1f1e1e",
    },
  ];
  return (
    <div className="h-[450px] w-full lg:w-1/3 flex flex-col bg-primary-light rounded-xl p-3 lg:p-4">
      <div className="w-full justify-between flex items-center">
        <h1 className="font-bold text-lg">Students</h1>
        <Image src="/more.png" alt="more" width={20} height={20} />
      </div>
      <div className="relative flex justify-center items-center w-full h-3/4">
        <RadialCharts data={data} />
        <div className="absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 flex gap-2 items-center">
          <Mars className="size-8 fill-accent-1 stroke-accent-1" />
          <Venus className="size-8 fill-accent-3 storke-accent-3" />
        </div>
      </div>
      <div className="flex w-full gap-2 justify-between items-center mt-4">
        <div className="flex text-start justify-center flex-col gap-1 text-sm">
          <p className="p-4 bg-sky rounded-full w-7 h-7" />
          <p className="font-semibold">{counts.maleStudents}</p>
          <p className="text-gray-400">
            Boys({((counts.maleStudents / counts.students) * 100).toFixed(2)}%)
          </p>
        </div>
        <div className="flex text-start justify-center flex-col gap-1 text-sm">
          <p className="p-4 bg-yellow rounded-full w-7 h-7" />
          <p className="font-semibold">{counts.femaleStudents}</p>
          <p className="text-gray-400">
            Girls({((counts.femaleStudents / counts.students) * 100).toFixed(2)}
            %)
          </p>
        </div>
      </div>
    </div>
  );
};

export default CountChart;
