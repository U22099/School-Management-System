"use client";

import Announcement from "@/components/Announcements";
import ScheduleCalendar from "@/components/ScheduleCalender";
import { useRole } from "@/store";
import { LoaderCircle } from "lucide-react";
import { useRouter } from "next/navigation";
import { useEffect } from "react";

const TeacherPage = () => {
  const router = useRouter();
  const role = useRole((state) => state.role);

  useEffect(() => {
    if (role !== "TEACHER") {
      router.push(`/${role.toLowerCase()}`);
    }
  }, []);
  if (role !== "TEACHER") {
    return (
      <div className="flex justify-center items-center w-full h-full gap-2 font-bold">
        <LoaderCircle className="animate-spin" /> You are not a TEACHER,
        redirecting to {role} page
      </div>
    );
  } else
    return (
      <>
        <section className="w-full flex gap-2 flex-col xl:flex-row h-fit">
          <div className="xl:w-2/3 flex flex-col gap-4 p-2 h-full">
            <div className="h-[850px] lg:h-full flex flex-col bg-white p-3 lg:p-4">
              <div className="w-full justify-between flex items-center">
                <h1 className="font-bold text-lg">Schedule</h1>
              </div>
              <ScheduleCalendar />
            </div>
          </div>
          <div className="xl:w-1/3 flex flex-col gap-4 p-2">
            <Announcement />
          </div>
        </section>
      </>
    );
};

export default TeacherPage;
