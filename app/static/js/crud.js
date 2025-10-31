function handleAjax(url, method, data, onSuccess, onError) {
  $.ajax({
    url, method,
    data: JSON.stringify(data || {}),
    contentType: 'application/json',
    dataType: 'json',
    headers: { 'Authorization': 'Bearer ' + "{{ session['jwt_token'] }}" },
    success: onSuccess,
    error: function(xhr, status, err){
      const payload = xhr.responseJSON || {error: xhr.responseText || err || 'Unknown error'};
      onError(payload, status, err);
    }
  });
}

function handleResponse(data, crud, reloadUrl){
  try{
    if (data && data.msg === "SUKSES"){
      alert(`✅ Sukses ${crud}`);
    } else {
      alert(`❌ Gagal ${crud}: ${(data && data.error) || 'unknown'}`);
    }
  }catch(e){
    alert('⚠ Error: '+e.message);
  }finally{
    setTimeout(()=>location.href = reloadUrl || location.href, 300);
  }
}

function showModal(title, html, onSubmit){
  $('#modal-title').text(title);
  $('#modal-body').html(html);
  $('#modal-submit').off('click').on('click', onSubmit);
  $('#generic-modal').modal('show');
}

// Build form fields per entity
function fieldsFor(entity){
  switch(entity){
    case 'products': return [
      {name:'code', label:'Kode', type:'text', required:true},
      {name:'name', label:'Nama', type:'text', required:true},
      {name:'unit', label:'Base Unit', type:'text', value:'pcs'},
      {name:'stock_min', label:'Stok Min', type:'number', value:0}
    ];
    case 'suppliers': return [
      {name:'name', label:'Nama', type:'text', required:true},
      {name:'address', label:'Alamat', type:'textarea'},
      {name:'phone', label:'Telepon', type:'text'}
    ];
    case 'customers': return [
      {name:'name', label:'Nama', type:'text', required:true},
      {name:'address', label:'Alamat', type:'textarea'},
      {name:'npwp', label:'NPWP', type:'text'}
    ];
    case 'salespersons': return [
      {name:'name', label:'Nama', type:'text', required:true},
      {name:'is_active', label:'Aktif', type:'checkbox', value:1}
    ];
    default: return [];
  }
}

function renderForm(fields, data){
  return fields.map(f=>{
    const v = (data && data[f.name] != null) ? data[f.name] : (f.value ?? '');
    if (f.type === 'textarea'){
      return `<div class="mb-2"><label class="form-label">${f.label}</label>
              <textarea class="form-control" id="f_${f.name}">${v??''}</textarea></div>`;
    }
    if (f.type === 'checkbox'){
      const checked = (String(v)=='1' || v===true) ? 'checked' : '';
      return `<div class="form-check mb-2">
        <input class="form-check-input" type="checkbox" id="f_${f.name}" ${checked}>
        <label class="form-check-label">${f.label}</label></div>`;
    }
    const req = f.required ? 'required' : '';
    const step = f.type==='number' ? ' step="1" ' : '';
    return `<div class="mb-2"><label class="form-label">${f.label}</label>
            <input class="form-control" id="f_${f.name}" type="${f.type}" ${req} ${step} value="${v??''}"></div>`;
  }).join('');
}

function collectData(fields){
  const d = {};
  fields.forEach(f=>{
    const el = document.getElementById('f_'+f.name);
    if (!el) return;
    if (f.type==='checkbox'){ d[f.name] = el.checked ? 1 : 0; }
    else if (f.type==='number'){ d[f.name] = Number(el.value||0); }
    else { d[f.name] = el.value ?? null; }
  });
  return d;
}

function openCreate(){
  const fields = fieldsFor(ENTITY);
  showModal(`Tambah ${ENTITY}`, renderForm(fields), function(){
    const payload = collectData(fields);
    handleAjax(`/admin/${ENTITY}`, 'POST', payload,
      (res)=>handleResponse(res, 'tambah', location.pathname),
      (err)=>handleResponse(err, 'tambah', location.pathname));
  });
}

function openEdit(id){
  // ambil baris dari tabel (DOM) supaya simpel
  // (atau buat endpoint GET detail jika mau)
  const tr = $(event.target).closest('tr')[0];
  const headers = Array.from(document.querySelectorAll('table thead th')).slice(1,-1);
  const tds = tr.querySelectorAll('td');
  const fields = fieldsFor(ENTITY);
  const data = { id: id };
  let idx=1;
  fields.forEach(f=>{
    data[f.name] = tds[idx++]?.innerText?.trim();
  });

  showModal(`Edit ${ENTITY} #${id}`, renderForm(fields, data), function(){
    const payload = { id: id, ...collectData(fields) };
    handleAjax(`/admin/${ENTITY}`, 'PUT', payload,
      (res)=>handleResponse(res, 'edit', location.pathname),
      (err)=>handleResponse(err, 'edit', location.pathname));
  });
}

function doDelete(id){
  if(!confirm('Yakin hapus?')) return;
  handleAjax(`/admin/${ENTITY}`, 'DELETE', {id},
    (res)=>handleResponse(res, 'hapus', location.pathname),
    (err)=>handleResponse(err, 'hapus', location.pathname));
}
