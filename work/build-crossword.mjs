import fs from "node:fs/promises";
import { Workbook, SpreadsheetFile } from "@oai/artifact-tool";
const csv = await fs.readFile("/Users/z3489864/Documents/Codex/2026-08-04/c/outputs/crossword-puzzle-template.csv","utf8");
const lines = csv.trim().split("\n").slice(1).map(line=>line.split(",").map(x=>x.replace(/^"|"$/g,"")));
const grid = Array.from({length:18},()=>Array(18).fill(""));
const nums = Array.from({length:18},()=>Array(18).fill(""));
for (const row of lines) { const cell=row[0], r=cell.charCodeAt(0)-65, c=Number(cell.slice(1))-1; grid[r][c]=row[3]||""; nums[r][c]=row[5]||""; }
const wb=Workbook.create();
const sheet=wb.worksheets.add("Crossword");
sheet.showGridLines=false;
sheet.getRange("A1:R18").values=grid;
sheet.getRange("A1:R18").format={font:{name:"Arial",size:16,bold:true,color:"#21332F"},horizontalAlignment:"center",verticalAlignment:"center",borders:{preset:"all",style:"thin",color:"#21332F"}};
sheet.getRange("A1:R18").format.columnWidth=4;
sheet.getRange("A1:R18").format.rowHeight=26;
for(let r=0;r<18;r++)for(let c=0;c<18;c++){const cell=sheet.getCell(r,c);if(!grid[r][c])cell.format={fill:"#21332F",font:{color:"#FFFFFF"}};else if(nums[r][c])cell.values=[[nums[r][c]+"\n"+grid[r][c]]];}
sheet.getRange("A1:R18").format.wrapText=true;
const clues=wb.worksheets.add("Clues");
clues.getRange("A1:E1").values=[["Direction","Number","Start cell","Clue","Answer"]];
clues.getRange("A1:E1").format={font:{bold:true,color:"#FFFFFF"},fill:"#3F7664"};
clues.getRange("A2:E7").values=[
["Across",1,"A2","Blitzkrieg . . .","BOP"],
["Across",3,"A6","What they paved to put up a parking lot.","PARADISE"],
["Across",6,"A15","Spirit in the ...","SKY"],
["Down",7,"B1","Tom T. Hall wrote this book.","WUTHERING HEIGHTS"],
["Across",18,"A18","The Stones suggest that if you try sometimes, you just might find you. . .","GET WHAT YOU NEED"],
["Across",19,"A19","Don't bring me down is the final track on this band's 1979 album Discovery","ELO"]];
clues.getRange("A1:E7").format.borders={preset:"all",style:"thin",color:"#D8E5DE"};
clues.getRange("A:E").format.columnWidth=22;
await fs.mkdir("/Users/z3489864/Documents/Codex/2026-08-04/c/outputs",{recursive:true});
const out=await SpreadsheetFile.exportXlsx(wb); await out.save("/Users/z3489864/Documents/Codex/2026-08-04/c/outputs/crossword-grid.xlsx");

