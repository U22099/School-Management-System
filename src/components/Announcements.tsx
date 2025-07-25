"use client";
import apiClient from "@/lib/apiclient";
import { useRecents } from "@/store";
import { Announcement as AnnouncementType } from "@prisma/client";
import Link from "next/link";
import { useEffect, useState } from "react";

const Announcement = () => {
  const { recents, setRecents } = useRecents();
  const [announcements, setAnnouncements] = useState<AnnouncementType[]>([]);

  useEffect(() => {
    const fetchAnnouncements = async () => {
      try {
        const res = await apiClient.get(
          "/announcements?page=1&order=desc&limit=3"
        );
        if (res.status === 200) {
          setAnnouncements(res.data.announcements);
          setRecents({
            announcements: res.data.announcements.filter(
              (x: AnnouncementType) =>
                new Date(x.date).toDateString() ===
                new Date(Date.now()).toDateString()
            ).length,
            events: recents.events,
          });
        }
      } catch (err) {
        console.log(err);
      }
    };

    fetchAnnouncements();
  }, []);
  return (
    <div className="flex flex-col gap-4 p-2 bg-primary-light w-full rounded-md">
      <div className="w-full justify-between flex items-center">
        <h1 className="font-bold text-lg">Announcements</h1>
        <Link className="text-sm text-gray-400" href="/list/announcements">
          View all
        </Link>
      </div>
      {!announcements.length && (
        <p className="font-bold text-gray-400 text-lg">No Announcements</p>
      )}
      {announcements.map((announcement, index) => (
        <div
          key={index}
          className="w-full p-3 lg:p-5 rounded-md h-fit bg-primary gap-2 border-2 odd:border-sky-light even:border-purple-light"
        >
          <div className="w-full justify-between flex items-center">
            <h3 className="font-semibold text-lg">{announcement.title}</h3>
            <p className="text-sm text-gray-400">
              {new Date(announcement.date).toLocaleDateString("en-UK", {
                day: "2-digit",
                month: "2-digit",
                year: "2-digit",
              })}
            </p>
          </div>
          <p className="text-gray-500">{announcement.description}</p>
        </div>
      ))}
    </div>
  );
};

export default Announcement;
