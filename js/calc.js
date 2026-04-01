function calcWireless(data){
  const result = [];
  for(let i=0;i<12;i++){
    const capa = data.kpi.capa[i];
    const active = data.kpi.active[i];

    const prod_rev = capa * data.unit.pu;
    const pol_fee = capa * data.unit.policy;
    const mgmt_fee = active * 3000;
    const sc_fee = capa * data.unit.sc;

    let promo = 0;
    if(i === 8) promo = capa * 20000;
    else if(i === 6) promo = capa * 15000;
    else if(i === 11) promo = capa * 30000;
    else promo = capa * 8000;

    result.push({month:i+1,prod_rev,pol_fee,mgmt_fee,sc_fee,promo});
  }
  return result;
}
