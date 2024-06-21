import { cn } from "@/lib/utils";
import Image from "next/image";
import { HTMLAttributes } from "react";

interface BulletinProps extends HTMLAttributes<HTMLDivElement> {
  imgSrc: string;
  dark?: boolean;
}

const Bulletin = ({ imgSrc, className, dark = false, ...props }: BulletinProps) => {
  return (
    <div className={cn("relative pointer-events-none z-50 overflow-hidden", className)} {...props}>
      <Image
        src={dark ? "/bulletin.png" : "/bulletin.png"}
        className="pointer-events-none z-50 select-none"
        alt="bulletin"
        width={956}
        height={956}
      />
    </div>
  );
};

export default Bulletin;
