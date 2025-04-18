type SchoolInfo = Omit<
  SchoolDataType,
  "admins" | "timetableHtml" | "subjects" | "id" | "schoolId"
>;
type StaffData = Omit<StaffDataType, "id" | "schoolId">;
type StudentData = Omit<StudentDataType, "id" | "schoolId">;
type ParentData = Omit<ParentDataType, "id" | "schoolId">;
type ClassData = Omit<ClassDataType, "id" | "schoolId">;
type Timetable = Omit<TimetableType, "id" | "schoolId">;
type Subject = string;

type TotalData = {
  schoolData: SchoolInfo | {};
  staffsData: StaffData[] | [];
  admins: StaffData[] | [];
  studentsData: StudentData[] | [];
  parentsData: ParentData[] | [];
  subjects: Subject[] | [];
  classes: ClassData[] | [];
  timetable: Timetable[] | [];
  timetableHtml: string;
};
