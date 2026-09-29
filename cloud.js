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

  async function signIn(email,password) {
    const c=init(); if(!c) throw new Error("Supabase aún no está configurado.");
    const {data,error}=await c.auth.signInWithPassword({email,password});
    if(error) throw error;
    await loadProfile();
    return data.session;
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
    const {data,error}=await c.from("profiles").select("id,email,nombre,role,activo").eq("id",session.user.id).single();
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

  async function listUsers() {
    const c=init(); if(!c) return [];
    const {data,error}=await c.from("profiles").select("id,email,nombre,role,activo,updated_at").order("nombre");
    if(error) throw error;
    return data||[];
  }

  async function setUserRole(userId,role) {
    const c=init(); if(!c) throw new Error("Supabase aún no está configurado.");
    const {error}=await c.rpc("set_user_role",{p_user:userId,p_role:role});
    if(error) throw error;
  }

  async function audit(limit=100) {
    const c=init(); if(!c) return [];
    const {data,error}=await c.from("audit_events").select("*").order("created_at",{ascending:false}).limit(limit);
    if(error) throw error;
    return data||[];
  }

  const role = () => profile?.role || "CONSULTA";
  const canEdit = () => ["ADMIN_TIC","EDITOR"].includes(role());
  const isAdmin = () => role()==="ADMIN_TIC";
  const getProfile = () => profile;
  const getVersion = () => version;

  return {configured,init,getSession,signIn,signOut,loadProfile,loadState,saveState,listUsers,setUserRole,audit,role,canEdit,isAdmin,getProfile,getVersion};
})();