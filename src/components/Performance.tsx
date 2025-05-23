"use client";

import Image from "next/image";
import { PieChart, Pie, ResponsiveContainer } from "recharts";

const Performance = ({ level }: { level: number }) => {
  const data = [
    { name: "Group A", value: level || 1, fill: "#C3EBFA" },
    { name: "Group B", value: (17 - level), fill: "#FAE27C" },
  ];
  return (
    <div className="bg-primary-light p-4 rounded-md h-80 relative">
      <div className="flex items-center justify-between">
        <h1 className="text-xl font-semibold">Level (1-17)</h1>
        <Image src="/moreDark.png" alt="" width={16} height={16} />
      </div>
      <ResponsiveContainer width="100%" height="100%">
        <PieChart>
          <Pie
            dataKey="value"
            startAngle={180}
            endAngle={0}
            data={data}
            cx="50%"
            cy="50%"
            innerRadius={70}
            fill="#8884d8"
          />
        </PieChart>
      </ResponsiveContainer>
      <div className="absolute top-1/2 left-1/2 transform -translate-x-1/2 -translate-y-1/2 text-center">
        <h1 className="text-3xl font-bold">{level}</h1>
        <p className="text-xs text-gray-300">of 17 max Level</p>
      </div>
      <h2 className="font-medium absolute bottom-16 left-0 right-0 m-auto text-center">
        {new Date().getFullYear()}
      </h2>
    </div>
  );
};

export default Performance;
