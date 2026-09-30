const PraxisCloud = (() => {
  let client = null;
  let profile = null;
  let version = 0;

  const cfg = () => window.PraxisSupabaseConfig || {};
  const configured = () => Boolean(cfg().url && cfg().anonKey && window.supabase?.createClient);

  function init() {
    if (!configured()) return null;
    if (!client) client = window.supabase.createClient(cfg().url, cfg().anonKey, {
      auth: { persistSession:true, autoRefreshToken:true, detectSessionInUrl:true }
    });
    return client;
  }

  async function getSession() {
    const c=init(); if(!c) return null;
    const {data,error}=await c.auth.getSession();
    if(error) throw error;
    return data.session;
  }

  function normalizeInstitutionalEmail(value) {
    const raw=String(value||"").trim().toLowerCase();
    if(!raw) return "";
    return raw.includes("@") ? raw : raw+"@praxis.edu.pe";
  }

  async function signIn(identifier,password) {
    const c=init(); if(!c) throw new Error("Supabase aún no está configurado.");
    const raw=String(identifier||"").trim();

    let session=null;

    if(raw==="71599459"){
      const {data,error}=await c.auth.signInWithPassword({
        email:"r.cardenas@praxis.edu.pe",
        password
      });
      if(error) throw error;
      session=data.session;
    } else if(/^\d{8}$/.test(raw)){
      const res=await fetch(cfg().url+"/functions/v1/login-dni",{
        method:"POST",
        headers:{
          "Content-Type":"application/json",
          "apikey":cfg().anonKey
        },
        body:JSON.stringify({dni:raw,password})
      });
      const payload=await res.json().catch(()=>({}));
      if(!res.ok || !payload?.access_token || !payload?.refresh_token){
        throw new Error("INVALID_LOGIN");
      }
      const {data,error}=await c.auth.setSession({
        access_token:payload.access_token,
        refresh_token:payload.refresh_token
      });
      if(error) throw error;
      session=data.session;
    } else {
      const email=normalizeInstitutionalEmail(raw);
      const {data,error}=await c.auth.signInWithPassword({email,password});
      if(error) throw error;
      session=data.session;
    }

    await loadProfile();
    try { await c.rpc("touch_login"); } catch(err) { console.warn("No se pudo registrar último acceso",err); }
    return session;
  }

  async function bootstrapAdmin() {
    const c=init(); if(!c) return false;
    const {data,error}=await c.rpc("bootstrap_first_admin");
    if(error) throw error;
    return Boolean(data);
  }

  async function signOut() {
    const c=init(); if(!c) return;
    const {error}=await c.auth.signOut();
    if(error) throw error;
    profile=null; version=0;
  }

  async function loadProfile() {
    const c=init(); if(!c) return null;
    const session=await getSession();
    if(!session){profile=null;return null;}
    const {data,error}=await c.from("profiles").select("id,email,nombre,dni,cargo,sede,area,role,activo,last_login_at,created_at,updated_at").eq("id",session.user.id).single();
    if(error) throw error;
    profile=data;
    return profile;
  }

  async function loadState() {
    const c=init(); if(!c) return {state:null,version:0};
    const {data,error}=await c.from("inventory_state").select("state,version,updated_at,updated_by").eq("id",1).single();
    if(error) throw error;
    version=Number(data?.version||0);
    return {state:data?.state||null,version,updatedAt:data?.updated_at,updatedBy:data?.updated_by};
  }

  async function saveState(state,action="ACTUALIZAR_INVENTARIO") {
    const c=init(); if(!c) throw new Error("Supabase aún no está configurado.");
    const {data,error}=await c.rpc("save_inventory_state",{
      p_payload:state,
      p_action:action,
      p_expected_version:version
    });
    if(error) {
      if(String(error.message||"").includes("VERSION_CONFLICT")) {
        const e=new Error("VERSION_CONFLICT");
        e.cause=error;
        throw e;
      }
      throw error;
    }
    const row=Array.isArray(data)?data[0]:data;
    version=Number(row?.new_version ?? version+1);
    return row;
  }

  async function prepareUser(payload) {
    const c=init(); if(!c) throw new Error("Supabase aún no está configurado.");
    const {data,error}=await c.rpc("prepare_inventory_user",{
      p_dni:String(payload?.dni||"").replace(/\D/g,""),
      p_nombre:String(payload?.nombre||"").trim(),
      p_cargo:String(payload?.cargo||"").trim(),
      p_sede:String(payload?.sede||"").trim(),
      p_area:String(payload?.area||"").trim(),
      p_role:String(payload?.role||"CONSULTA").toUpperCase()
    });
    if(error){
      const msg=String(error.message||"");
      if(msg.includes("DNI_ALREADY_EXISTS")) throw new Error("Ya existe un usuario con ese DNI.");
      if(msg.includes("INVALID_DNI")) throw new Error("El DNI debe tener 8 dígitos.");
      if(msg.includes("ROLE_NOT_ALLOWED")) throw new Error("Solo puedes asignar rol Editor o Consulta.");
      if(msg.includes("OWNER_ADMIN_REQUIRED")) throw new Error("Solo la cuenta Master puede preparar usuarios.");
      if(msg.includes("STAFF_DATA_REQUIRED")) throw new Error("Completa cargo, sede y área.");
      throw error;
    }
    return String(data||"");
  }

  async function listUsers() {
    const c=init(); if(!c) return [];
    const {data,error}=await c.rpc("admin_list_users");
    if(error) throw error;
    return data||[];
  }

  async function setUserRole(userId,role) {
    const c=init(); if(!c) throw new Error("Supabase aún no está configurado.");
    const {error}=await c.rpc("set_user_role",{p_user:userId,p_role:role});
    if(error) throw error;
  }

  async function setUserActive(userId,active) {
    const c=init(); if(!c) throw new Error("Supabase aún no está configurado.");
    const {error}=await c.rpc("set_user_active",{p_user:userId,p_active:Boolean(active)});
    if(error) throw error;
  }


  async function audit(limit=100) {
    const c=init(); if(!c) return [];
    const {data,error}=await c.from("audit_events").select("*").order("created_at",{ascending:false}).limit(limit);
    if(error) throw error;
    return data||[];
  }

  function subscribeProfile(onProfileChange) {
    const c=init();
    if(!c || !profile?.id) return null;
    const channel=c.channel("praxis-profile-"+profile.id)
      .on("postgres_changes",{event:"UPDATE",schema:"public",table:"profiles",filter:"id=eq."+profile.id},async()=>{
        try{
          const latest=await loadProfile();
          if(typeof onProfileChange==="function") onProfileChange(latest);
        }catch(err){ console.error("Realtime profile reload failed",err); }
      })
      .subscribe();
    return channel;
  }

  function subscribeState(onRemoteState) {
    const c=init();
    if(!c) return null;
    const channel=c.channel("praxis-inventory-state")
      .on("postgres_changes",{event:"UPDATE",schema:"public",table:"inventory_sync",filter:"id=eq.1"},async payload=>{
        const remoteVersion=Number(payload?.new?.version||0);
        if(remoteVersion<=version) return;
        try{
          const latest=await loadState();
          if(typeof onRemoteState==="function") onRemoteState(latest);
        }catch(err){ console.error("Realtime inventory reload failed",err); }
      })
      .subscribe();
    return channel;
  }

  async function unsubscribe(channel) {
    const c=init();
    if(c && channel) await c.removeChannel(channel);
  }

  const role = () => profile?.role || "CONSULTA";
  const canEdit = () => ["ADMIN_TIC","EDITOR"].includes(role());
  const isAdmin = () => role()==="ADMIN_TIC";
  const getProfile = () => profile;
  const getVersion = () => version;

  return {configured,init,getSession,normalizeInstitutionalEmail,signIn,bootstrapAdmin,signOut,loadProfile,loadState,saveState,prepareUser,listUsers,setUserRole,setUserActive,audit,subscribeProfile,subscribeState,unsubscribe,role,canEdit,isAdmin,getProfile,getVersion};
})();