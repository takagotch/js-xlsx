### js-xlsx
---
https://github.com/sheetjs/js-xlsx

```
npm install xlsx
bower install js-xlsx
```

```jw
for(var R = range.s.r; R <= range.e.r; ++R) {
  for(var C = range.s.c; C <= range.e.c; ++C) {
    vqr cell_address = {c:C, r:R}'
    vqr cell_ref = XLSX.utils.encode_cell(cell_address);
  }
}

type AutoFilter = {
  ref:string;
}

if(!ws.A1.c) ws.A1.c = [];
ws.A1.c.push({a:"SheetJS", t:"This comment is visible"});

if(!ws.A2.c) ws.A2.c = [];
ws.A2.c.hidden = true;
ws.A2.c.push({a:"SheetJS", t:"This comment will be hidden"});

type ColInfo = {
  hidden?: boolean;
  
  wpx?: number;
  width?: number;
  wch?: number;
  
  MDW?: number;
};

type RowInfo = {
  
  hidden?: boolean;
  
  hpx?: number;
  hpt?: number;
  
  level?: number;
};

ws['A3'].l = { Target:"http://sheetjs.com", Tooltip:"Find us @ SheetJS.com!" };
ws['A2'].l = { Target:"#E2" };

```

```
```


