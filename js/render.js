function fmt(n){
  return Math.round(n).toLocaleString();
}

function renderTable(data){
  const tbody = document.getElementById("wireless-body");
  tbody.innerHTML = "";

  data.forEach(row=>{
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${row.month}</td>
      <td>${fmt(row.prod_rev)}</td>
      <td>${fmt(row.pol_fee)}</td>
      <td>${fmt(row.mgmt_fee)}</td>
      <td>${fmt(row.sc_fee)}</td>
      <td>${fmt(row.promo)}</td>
    `;
    tbody.appendChild(tr);
  });
}
