// function showTable(data) {
//   const table = document.getElementById("dataTable");
//   table.innerHTML = "";

//   if (data.length === 0) {
//     table.innerHTML = "<tr><td>No transactions found</td></tr>";
//     return;
//   }

//   const headers = Object.keys(data[0]);

//   const headerRow = document.createElement("tr");

//   headers.forEach(header => {
//     const th = document.createElement("th");
//     th.innerText = header;
//     headerRow.appendChild(th);
//   });

//   table.appendChild(headerRow);

//   data.forEach(row => {
//     const tr = document.createElement("tr");

//     headers.forEach(header => {
//       const td = document.createElement("td");
//       td.innerText = row[header] ?? "";
//       tr.appendChild(td);
//     });

//     table.appendChild(tr);
//   });
// }