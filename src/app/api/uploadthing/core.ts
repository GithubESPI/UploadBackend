import { db } from "@/db";
import ExcelJS from "exceljs";
import fetch from "node-fetch";
import { createUploadthing, type FileRouter } from "uploadthing/next";
import { z } from "zod";

const f = createUploadthing();

export const ourFileRouter = {
  fileUploader: f({
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": { maxFileSize: "4GB" },
    "application/vnd.ms-excel": { maxFileSize: "4GB" },
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document": {
      maxFileSize: "4GB",
    },
    "application/msword": { maxFileSize: "4GB" },
  })
    .input(z.object({ configId: z.string().optional() }))
    .middleware(async ({ input }) => {
      return { input };
    })
    .onUploadComplete(async ({ metadata, file }) => {
      const { configId } = metadata.input;

      try {
        // Fetch the file and read it as a buffer
        const res = await fetch(file.url);
        const buffer = await res.arrayBuffer();

        // Use exceljs to parse the Excel file
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(buffer);

        const worksheet = workbook.getWorksheet(1);
        const excelData = [];

        if (worksheet) {
          worksheet.eachRow((row, rowNumber) => {
            excelData.push(row.values);
          });
        } else {
          throw new Error("Worksheet not found");
        }

        const excelUrl = file.url; // Get the URL of the uploaded file
        const fileName = file.name; // Get the name of the uploaded file

        if (!configId) {
          const configuration = await db.configuration.create({
            data: {
              excelUrl: excelUrl,
              fileName: fileName,
            },
          });

          return { configId: configuration.id };
        } else {
          const updatedConfiguration = await db.configuration.update({
            where: {
              id: configId,
            },
            data: {
              excelUrl: excelUrl,
              fileName: fileName,
            },
          });

          return { configId: updatedConfiguration.id };
        }
      } catch (error) {
        console.error("Error in onUploadComplete:", error);
        throw new Error("Failed to process the uploaded file");
      }
    }),
} satisfies FileRouter;

export type OurFileRouter = typeof ourFileRouter;
