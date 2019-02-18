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
```

```
```


