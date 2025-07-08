"use client";

const Pagination = ({
  currentPage,
  totalPages,
  onPageChange,
}: {
  currentPage: number;
  totalPages: number;
  onPageChange: (page: number) => void;
}) => {
  return (
    <div className="p-4 flex items-center justify-between text-primary">
      <button
        onClick={currentPage !== 1 ? () => onPageChange(currentPage - 1) : () => console.log("Nice Try")}
        disabled={currentPage === 1}
        className="py-2 px-4 rounded-md bg-secondary text-xs font-semibold disabled:opacity-50 disabled:cursor-not-allowed"
      >
        Prev
      </button>
      <div className="flex items-center gap-2 text-sm">
        {Array.from({ length: totalPages }, (_, i) => i + 1).map((page) => (
          <button
            key={page}
            onClick={() => onPageChange(page)}
            className={`px-2 rounded-sm ${
              page === currentPage ? "bg-accent-1" : ""
            }`}
          >
            {page}
          </button>
        ))}
      </div>
      <button
        onClick={currentPage !== totalPages ? () => onPageChange(currentPage + 1) : () => console.log("Nice Try")}
        disabled={currentPage === totalPages}
        className="py-2 px-4 rounded-md bg-secondary text-xs font-semibold disabled:opacity-50 disabled:cursor-not-allowed"
      >
        Next
      </button>
    </div>
  );
};

export default Pagination;