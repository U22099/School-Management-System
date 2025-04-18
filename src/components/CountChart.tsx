"use client";

import Image from "next/image";
import RadialCharts from "./RadialChart";

const CountChart = () => {
  const data = [
    {
      name: "Male",
      count: 1334,
      fill: "#ceebfa",
    },
    {
      name: "Female",
      count: 1234,
      fill: "#fae27c",
    },
    {
      name: "Gender",
      count: 1234 * 2 + 100,
      fill: "white",
    },
  ];
  return (
    <div className="h-[450px] w-full lg:w-1/3 flex flex-col bg-white rounded-xl p-3 lg:p-4">
      <div className="w-full justify-between flex items-center">
        <h1 className="font-bold text-lg">Students</h1>
        <Image src="/moreDark.png" alt="more" width={20} height={20} />
      </div>
      <div className="relative flex justify-center items-center w-full h-3/4">
        <RadialCharts data={data} />
        <Image
          src="/maleFemale.png"
          alt="maleFemale"
          width={50}
          height={50}
          className="absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2"
        />
      </div>
      <div className="flex w-full gap-2 justify-between items-center mt-4">
        <div className="flex text-start justify-center flex-col gap-1 text-sm">
          <p className="p-4 bg-sky rounded-full w-7 h-7" />
          <p className="font-semibold">1,234</p>
          <p className="text-gray-400">Boys({55}%)</p>
        </div>
        <div className="flex text-start justify-center flex-col gap-1 text-sm">
          <p className="p-4 bg-yellow rounded-full w-7 h-7" />
          <p className="font-semibold">1,234</p>
          <p className="text-gray-400">Girls({55}%)</p>
        </div>
      </div>
    </div>
  );
};

export default CountChart;
