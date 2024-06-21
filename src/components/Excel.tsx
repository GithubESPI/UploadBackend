import { cn } from "@/lib/utils";
import Image from "next/image";
import { HTMLAttributes } from "react";

interface ExcelProps extends HTMLAttributes<HTMLDivElement> {
  imgSrc: string;
  dark?: boolean;
}

const Excel = ({ imgSrc, className, dark = false, ...props }: ExcelProps) => {
  return (
    <div className={cn("relative pointer-events-none z-50 overflow-hidden", className)} {...props}>
      <Image
        src={dark ? "/img-3.png" : "/img-3.png"}
        className="pointer-events-none z-50 select-none"
        alt="darkexcel"
        width={956}
        height={956}
      />
    </div>
  );
};

export default Excel;
