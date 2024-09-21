"use client";

import React, { useState } from "react";
import * as XLSX from "xlsx"; // xlsx-style 라이브러리 임포트
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { PlusCircle, Download, Trash2 } from "lucide-react";

interface Item {
  순번: number;
  매입처: string;
  품목: string;
  규격: string;
  원산지: string;
  수량: number;
  매입단가: number;
  매출단가: number;
}

export function ExcelInputAppComponent() {
  const [items, setItems] = useState<Item[]>([
    {
      순번: 1,
      매입처: "",
      품목: "",
      규격: "",
      원산지: "",
      수량: 0,
      매입단가: 0,
      매출단가: 0,
    },
  ]);

  const handleChange = (
    index: number,
    field: keyof Item,
    value: string | number
  ) => {
    const newItems = [...items];
    newItems[index][field] = value as never;
    setItems(newItems);
  };

  const addItem = () => {
    setItems([
      ...items,
      {
        순번: items.length + 1,
        매입처: "",
        품목: "",
        규격: "",
        원산지: "",
        수량: 0,
        매입단가: 0,
        매출단가: 0,
      },
    ]);
  };

  const removeItem = (index: number) => {
    const newItems = items.filter((_, i) => i !== index);
    setItems(newItems);
  };

  const exportToExcel = async () => {
    // 샘플 xlsx 파일 읽기
    const response = await fetch("/template.xlsx");
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });

    // 첫 번째 시트 선택
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];

    items.forEach((item, index) => {
      const rowIndex = index + 5; // A5부터 시작
      worksheet[`A${rowIndex}`] = {
        t: "n",
        v: item.순번,
        s: worksheet[`A${rowIndex}`]?.s,
      };
      worksheet[`B${rowIndex}`] = {
        t: "s",
        v: item.매입처,
        s: worksheet[`B${rowIndex}`]?.s,
      };
      worksheet[`C${rowIndex}`] = {
        t: "s",
        v: item.품목,
        s: worksheet[`C${rowIndex}`]?.s,
      };
      worksheet[`D${rowIndex}`] = {
        t: "s",
        v: item.규격,
        s: worksheet[`D${rowIndex}`]?.s,
      };
      worksheet[`E${rowIndex}`] = {
        t: "s",
        v: item.원산지,
        s: worksheet[`E${rowIndex}`]?.s,
      };
      worksheet[`F${rowIndex}`] = {
        t: "n",
        v: item.수량,
        s: worksheet[`F${rowIndex}`]?.s,
      };
      worksheet[`G${rowIndex}`] = {
        t: "n",
        v: item.매입단가,
        s: worksheet[`G${rowIndex}`]?.s,
      };
      worksheet[`G${rowIndex}`] = {
        t: "n",
        v: item.수량 * item.매입단가,
        s: worksheet[`H${rowIndex}`]?.s,
      };
      worksheet[`I${rowIndex}`] = {
        t: "n",
        v: item.매출단가,
        s: worksheet[`H${rowIndex}`]?.s,
      };
      worksheet[`J${rowIndex}`] = {
        t: "n",
        v: item.수량 * item.매출단가,
        s: worksheet[`H${rowIndex}`]?.s,
      };
      worksheet[`K${rowIndex}`] = {
        t: "n",
        v: item.매출단가 - item.매입단가,
        s: worksheet[`H${rowIndex}`]?.s,
      };
      worksheet[`L${rowIndex}`] = {
        t: "n",
        v: (item.매출단가 - item.매입단가) * item.수량,
        s: worksheet[`H${rowIndex}`]?.s,
      };
    });

    // 엑셀 파일 저장
    const dateStr = new Date().toISOString().slice(0, 10); // yyyy-MM-dd 형식
    XLSX.writeFile(workbook, `${dateStr} 매입 매출 명세서.xlsx`);
  };
  return (
    <div className="p-4 max-w-md mx-auto">
      <h1 className="text-2xl font-bold mb-4 text-center">항목 입력</h1>
      {items.map((item, index) => (
        <div key={index} className="mb-6 p-4 border rounded-lg shadow-sm">
          <div className="grid grid-cols-2 gap-4">
            <div>
              <Label htmlFor={`supplier-${index}`}>매입처</Label>
              <Input
                id={`supplier-${index}`}
                value={item.매입처}
                onChange={(e) => handleChange(index, "매입처", e.target.value)}
                className="mt-1"
              />
            </div>
            <div>
              <Label htmlFor={`product-${index}`}>품목</Label>
              <Input
                id={`product-${index}`}
                value={item.품목}
                onChange={(e) => handleChange(index, "품목", e.target.value)}
                className="mt-1"
              />
            </div>
            <div>
              <Label htmlFor={`specification-${index}`}>규격</Label>
              <Input
                id={`specification-${index}`}
                value={item.규격}
                onChange={(e) => handleChange(index, "규격", e.target.value)}
                className="mt-1"
              />
            </div>
            <div>
              <Label htmlFor={`origin-${index}`}>원산지</Label>
              <Input
                id={`origin-${index}`}
                value={item.원산지}
                onChange={(e) => handleChange(index, "원산지", e.target.value)}
                className="mt-1"
              />
            </div>
            <div>
              <Label htmlFor={`quantity-${index}`}>수량</Label>
              <Input
                id={`quantity-${index}`}
                value={item.수량}
                onChange={(e) => handleChange(index, "수량", e.target.value)}
                className="mt-1"
              />
            </div>
            <div>
              <Label htmlFor={`purchasePrice-${index}`}>매입 단가</Label>
              <Input
                id={`purchasePrice-${index}`}
                value={item.매입단가}
                onChange={(e) =>
                  handleChange(index, "매입단가", e.target.value)
                }
                className="mt-1"
              />
            </div>
            <div>
              <Label htmlFor={`sellingPrice-${index}`}>매출 단가</Label>
              <Input
                id={`sellingPrice-${index}`}
                value={item.매출단가}
                onChange={(e) =>
                  handleChange(index, "매출단가", e.target.value)
                }
                className="mt-1"
              />
            </div>
          </div>
          <Button
            variant="destructive"
            size="icon"
            className="mt-4"
            onClick={() => removeItem(index)}
          >
            <Trash2 className="h-4 w-4" />
          </Button>
        </div>
      ))}
      <div className="flex justify-between mt-4">
        <Button onClick={addItem}>
          <PlusCircle className="mr-2 h-4 w-4" /> 항목 추가
        </Button>
        <Button onClick={exportToExcel}>
          <Download className="mr-2 h-4 w-4" /> 엑셀로 내보내기
        </Button>
      </div>
    </div>
  );
}
