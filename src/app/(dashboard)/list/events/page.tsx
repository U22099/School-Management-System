"use client";

import FormModal from "@/components/FormModal";
import Pagination from "@/components/Pagination";
import Table from "@/components/Table";
import TableSearch from "@/components/TableSearch";
import apiClient from "@/lib/apiclient";
import { useRole } from "@/store";
import { Event } from "@prisma/client";
import Image from "next/image";
import { useEffect, useState } from "react";
import toast from "react-hot-toast";

const columns = [
  {
    header: "Title",
  },
  {
    header: "Date",
    className: "hidden md:table-cell",
  },
  {
    header: "Start Time",
    className: "hidden md:table-cell",
  },
  {
    header: "End Time",
    className: "hidden md:table-cell",
  },
  {
    header: "Actions",
  },
];

const EventListPage = () => {
  const role = useRole((state) => state.role);
  const [events, setEvents] = useState([]);
  const [page, setPage] = useState(1);
  const [totalPages, setTotalPages] = useState(1);

  const fetchEvents = async (page: number) => {
    try {
      const res = await apiClient.get(`/events?page=${page}&limit=10`);
      if (res.status === 200) {
        setEvents(res.data.events);
        setTotalPages(res.data.totalPages);
        toast.dismiss();
      } else {
        toast.dismiss();
        toast.error(res.data.message || "Failed to fetch events");
      }
    } catch (error) {
      toast.dismiss();
      console.error("Error fetching events:", error);
      toast.error("An error occurred while fetching events");
    }
  };

  const renderRow = (item: Event) => (
    <tr
      key={item.id}
      className="border-b border-gray-200 even:bg-primary-light text-sm hover:cursor-pointer"
    >
      <td className="flex items-center gap-4 p-4">{item.title}</td>
      <td className="hidden md:table-cell">{new Date(item.date).toDateString()}</td>
      <td className="hidden md:table-cell">{item.startTime}</td>
      <td className="hidden md:table-cell">{item.endTime}</td>
      <td>
        <div className="flex items-center gap-2">
          {role === "ADMIN" && (
            <>
              <FormModal table="events" type="update" data={item} />
              <FormModal table="events" type="delete" id={item.id} />
            </>
          )}
        </div>
      </td>
    </tr>
  );

  useEffect(() => {
    toast.loading("Fetching Data...");
    fetchEvents(page);
  }, [page]);

  return (
    <div className="bg-primary-light p-4 rounded-md flex-1 m-4 mt-0">
      {/* TOP */}
      <div className="flex items-center justify-between">
        <h1 className="hidden md:block text-lg font-semibold">All Events</h1>
        <div className="flex flex-col md:flex-row items-center gap-4 w-full md:w-auto">
          <TableSearch />
          <div className="flex items-center gap-4 self-end">
            <button className="w-8 h-8 flex items-center justify-center rounded-full bg-accent-3">
              <Image src="/filter.png" alt="" width={14} height={14} />
            </button>
            <button className="w-8 h-8 flex items-center justify-center rounded-full bg-accent-3">
              <Image src="/sort.png" alt="" width={14} height={14} />
            </button>
            {role === "ADMIN" && <FormModal table="events" type="create" />}
          </div>
        </div>
      </div>
      {/* LIST */}
      <Table columns={columns} renderRow={renderRow} data={events} />
      {/* PAGINATION */}
      <Pagination
        currentPage={page}
        totalPages={totalPages}
        onPageChange={(newPage) => setPage(newPage)}
      />
    </div>
  );
};

export default EventListPage;
