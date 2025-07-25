"use client";

import { useForm } from "react-hook-form";
import { z } from "zod";
import { zodResolver } from "@hookform/resolvers/zod";
import { ArrowLeft, ArrowRight, Upload } from "lucide-react";
import { motion } from "framer-motion";
import { useState } from "react";
import Image from "next/image";

const schema = z.object({
  name: z.string().min(3, { message: "School name is required!" }),
  address: z.string().min(3, { message: "Address is required!" }),
  principal: z.string().min(3, { message: "Principal name is required!" }),
  vicePrincipal: z
    .string()
    .min(3, { message: "Vice Principal name is required!" }),
  slogan: z.string().min(3, { message: "Slogan is required!" }),
  missionStatement: z
    .string()
    .min(3, { message: "Vision statement is required!" }),
  visionStatement: z
    .string()
    .min(3, { message: "Vision statement is required!" }),
  type: z.enum(["PRIMARY", "JUNIOR", "SENIOR"], {
    message: "School Type can either be PRIMARY, JUNIOR or SENIOR",
  }),
  startHour: z.string().min(4, { message: "Start Hour is required!" }),
  closeHour: z.string().min(4, { message: "Close Hour is required!" }),
  logo: z.string().min(1, { message: "Logo must be a valid Image!" }),
  primaryColor: z.string(),
  primaryColorLight: z.string(),
  secondaryColor: z.string(),
  secondaryColorLight: z.string(),
  accentColor1: z.string(),
  accentColor1Light: z.string(),
  accentColor2: z.string(),
  accentColor2Light: z.string(),
  accentColor3: z.string(),
  accentColor3Light: z.string(),
});

type FormData = z.infer<typeof schema>;

const SchoolInfoForm = ({
  setTotalData,
  setPage,
  defaultValues,
}: {
  setTotalData: React.Dispatch<React.SetStateAction<TotalData>>;
  setPage: React.Dispatch<React.SetStateAction<number>>;
  defaultValues: SchoolInfo | {};
}) => {
  const [logoBase64, setLogoBase64] = useState<string | null>(
    "logo" in defaultValues ? defaultValues.logo : null
  );

  const {
    register,
    handleSubmit,
    setValue,
    formState: { errors },
  } = useForm<FormData>({
    resolver: zodResolver(schema),
    mode: "onBlur",
    defaultValues,
  });

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = () => {
        const base64 = reader.result as string;
        setLogoBase64(base64);
        setValue("logo", base64);
      };
      reader.readAsDataURL(file);
    }
  };

  const onSubmit = async (data: FormData) => {
    try {
      console.log(data);
      setTotalData((prev) => ({ ...prev, schoolData: data }));
      setPage((prev) => prev + 1);
    } catch (error) {
      console.error("Error creating school:", error);
    }
  };

  return (
    <motion.div
      initial={{ opacity: 0 }}
      animate={{ opacity: 1 }}
      transition={{ duration: 0.6 }}
      className="flex items-center justify-center w-full h-full overflow-y-scroll"
    >
      <form
        onSubmit={handleSubmit(onSubmit)}
        className="p-6 rounded-lg shadow-lg w-full h-full flex flex-col gap-4 mb-4 mt-5"
      >
        <h1 className="text-2xl font-bold mb-4">Register Your School</h1>
        <div className="w-full flex-col flex items-between">
          <div className="gap-4 w-full md:flex-row md:flex-wrap flex-col flex justify-center">
            {Object.keys(schema.shape).map((key) => (
              <div className="min-w-[250px] flex-1 flex flex-col" key={key}>
                <label className="block text-sm font-medium capitalize">
                  School&apos;s {key.replace(/([A-Z])/g, " $1")}
                </label>
                {key === "logo" && (
                  <div className="flex gap-2 w-full items-center border-b border-primary rounded-md outline-none cursor-pointer">
                    <label
                      htmlFor={key}
                      className="flex gap-2 items-center w-full cursor-pointer p-2"
                    >
                      <Upload /> Upload Image
                    </label>
                    <input
                      type="file"
                      accept="image/*"
                      id={key}
                      hidden
                      onChange={handleFileChange}
                      className="input"
                    />
                    {logoBase64 && (
                      <Image
                        width={50}
                        height={50}
                        src={logoBase64}
                        alt="Preview"
                        className="rounded-md w-10 h-10"
                      />
                    )}
                  </div>
                )}
                {key === "type" ? (
                  <select
                    {...register("type")}
                    className="input p-2 border-b border-primary rounded-md outline-none w-full"
                    defaultValue=""
                  >
                    <option value="" disabled>
                      Select School Type
                    </option>
                    <option value="PRIMARY">Primary</option>
                    <option value="JUNIOR">Junior</option>
                    <option value="SENIOR">Senior</option>
                  </select>
                ) : key !== "logo" && (
                  <input
                    type={
                      key.includes("Color")
                        ? "color"
                        : key.includes("Hour")
                        ? "time"
                        : "text"
                    }
                    {...register(key as keyof FormData)}
                    className={`input p-2 border-b border-primary rounded-md outline-none ${
                      key.includes("Color") ? "h-10 w-full" : "w-full"
                    }`}
                  />
                )}
                {errors[key as keyof FormData] && (
                  <p className="text-red-500 text-sm mt-2">
                    {errors[key as keyof FormData]?.message}
                  </p>
                )}
              </div>
            ))}
          </div>
          <p className="text-xs font-semibold text-secondary-light mt-2">
            All field&apos;s are required except color palette as you can leave
            it default and change it later in the settings
          </p>
        </div>
        <div className="w-full flex gap-4 justify-between">
          <button
            disabled
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
                    (i === 0 ? "bg-secondary" : "bg-transparent")
                  }
                />
              ))}
          </div>
          <button type="submit" className="button justify-end">
            Next <ArrowRight className="stroke-black" />
          </button>
        </div>
      </form>
    </motion.div>
  );
};

export default SchoolInfoForm;
