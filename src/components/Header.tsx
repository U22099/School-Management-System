"use client";

import Image from "next/image";
import Link from "next/link";

// Code: Header component
const Header = () => {
  return (
    <main className="flex justify-end md:justify-between w-full p-3 items-center">
      <div className="hidden md:flex items-center gap-2 px-2 ring-2 ring-gray-300 rounded-full text-xs">
        <Image src="/search.png" width={14} height={14} alt="search" />
        <input
          type="text"
          placeholder="Search..."
          className="bg-transparent outline-none md:w-[200px] p-2"
        />
      </div>
      <div className="flex items-center justify-between gap-2">
        <div className="rounded-full p-2 bg-white cursor-pointer">
          <Image
            src="/announcement.png"
            width={20}
            height={20}
            alt="announcement"
          />
        </div>
        <div className="rounded-full p-2 bg-white cursor-pointer">
          <Image src="/profile.png" width={20} height={20} alt="profile" />
        </div>
        <div className="flex flex-col items-end justify-center gap-1 text-sm">
            <h1 className="font-bold">John Doe</h1>
            <p className="text-xs text-gray-400 capitalize">admin</p>
        </div>
        <Link href="/profile" className="rounded-full cursor-pointer overflow-hidden">
            <Image src="/avatar.png" width={50} height={50} alt="avatar" />
        </Link>
      </div>
    </main>
  );
};

export default Header;
