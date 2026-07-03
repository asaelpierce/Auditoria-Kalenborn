import React, { useState, useEffect, useRef, useCallback, useMemo } from 'react'
import { flushSync } from 'react-dom'
import {
  ClipboardCheck, FileText, AlertTriangle, Save, Plus, CheckCircle, XCircle,
  FolderTree, LayoutDashboard, Building2, Bell, Search, User, Printer, ChevronDown,
  Gauge, TrendingUp, ShieldCheck, Clock, Trash2, X, FileSpreadsheet, AlertCircle,
  ListPlus, Link2, Calendar, ChevronLeft, ChevronRight, Mail, GripVertical,
  CalendarDays, CalendarRange, Copy, Check, Folder, Upload, ExternalLink,
  FileWarning, FileBadge, Filter, Tag, Edit3
} from 'lucide-react'
import * as XLSX from 'xlsx'

const KALEN_LOGO_SRC        = '/logo-preta.png';
const KALEN_LOGO_BRANCA_SRC = '/logo-branca.png';

// ErrorBoundary — captura erros de render e mostra mensagem em vez de tela branca
class ErrorBoundary extends React.Component {
  constructor(props) { super(props); this.state = { error: null }; }
  static getDerivedStateFromError(error) { return { error }; }
  render() {
    if (this.state.error) {
      return (
        <div style={{ padding: 40, fontFamily: 'sans-serif', background: '#fff1f2', minHeight: '100vh' }}>
          <h2 style={{ color: '#be123c', marginBottom: 12 }}>⚠️ Erro na renderização</h2>
          <p style={{ color: '#881337', marginBottom: 8 }}>Copie a mensagem abaixo e envie para o suporte:</p>
          <pre style={{ background: '#ffe4e6', padding: 16, borderRadius: 8, fontSize: 13, whiteSpace: 'pre-wrap', wordBreak: 'break-all' }}>
            {this.state.error?.toString()}{'\n\n'}{this.state.error?.stack}
          </pre>
          <button onClick={() => this.setState({ error: null })} style={{ marginTop: 16, padding: '8px 20px', background: '#be123c', color: '#fff', border: 'none', borderRadius: 8, cursor: 'pointer', fontWeight: 'bold' }}>
            Tentar novamente
          </button>
        </div>
      );
    }
    return this.props.children;
  }
}

const KalenLogo = ({ variant = 'full', className = '' }) => {
  if (variant === 'sidebar') {
    return (
      <div className={`flex items-center ${className}`}>
        <img
          src={KALEN_LOGO_BRANCA_SRC}
          alt="Kalenborn Wear Protection Solutions"
          style={{ height: '36px', width: 'auto', objectFit: 'contain' }}
        />
      </div>
    );
  }
  // 'full' — formulários (logo preta sobre fundo branco)
  return (
    <div className={`flex items-center ${className}`}>
      <img
        src={KALEN_LOGO_SRC}
        alt="Kalenborn Wear Protection Solutions"
        style={{ height: '40px', width: 'auto', objectFit: 'contain', maxWidth: '180px' }}
      />
    </div>
  );
};

// ---------------------------------------------------------------------------
// SUPABASE — funções de acesso ao banco e storage
// ---------------------------------------------------------------------------
const SUPABASE_URL = 'https://zybkcpvdptabxkxpieuv.supabase.co';
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inp5YmtjcHZkcHRhYnhreHBpZXV2Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODEwMjU2MjUsImV4cCI6MjA5NjYwMTYyNX0.2BPIEHBuURfelci8YSVt7_PViJWMC1n853J2a3KC3HU';
const SGA_BUCKET = 'sga-documentos';

const supaFetch = async (path, opts = {}) => {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/${path}`, {
    headers: {
      'apikey': SUPABASE_ANON_KEY,
      'Authorization': `Bearer ${SUPABASE_ANON_KEY}`,
      'Content-Type': 'application/json',
      'Prefer': 'return=representation',
      ...opts.headers
    },
    ...opts
  });
  if (!res.ok) throw new Error(`Supabase ${res.status}: ${await res.text()}`);
  return res.json();
};

// Busca TODAS as linhas de uma consulta, paginando automaticamente (o Supabase
// limita a ~1000 linhas por requisição por padrão). Usado para agregados que
// precisam olhar o histórico inteiro (ex.: contagem de C/NC/Obs de todas as
// auditorias já feitas, não só a página mais recente).
const supaFetchAll = async (path, batchSize = 1000) => {
  let all = [];
  let offset = 0;
  // eslint-disable-next-line no-constant-condition
  while (true) {
    const sep = path.includes('?') ? '&' : '?';
    const res = await fetch(`${SUPABASE_URL}/rest/v1/${path}${sep}limit=${batchSize}&offset=${offset}`, {
      headers: {
        'apikey': SUPABASE_ANON_KEY,
        'Authorization': `Bearer ${SUPABASE_ANON_KEY}`,
      }
    });
    if (!res.ok) throw new Error(`Supabase ${res.status}: ${await res.text()}`);
    const batch = await res.json();
    all = all.concat(batch);
    if (batch.length < batchSize) break;
    offset += batchSize;
  }
  return all;
};

const supaStorageUpload = async (file) => {
  const safeName = file.name.replace(/[^a-zA-Z0-9._-]/g, '_');
  const path = `${Date.now()}_${safeName}`;
  const res = await fetch(`${SUPABASE_URL}/storage/v1/object/${SGA_BUCKET}/${path}`, {
    method: 'POST',
    headers: {
      'apikey': SUPABASE_ANON_KEY,
      'Authorization': `Bearer ${SUPABASE_ANON_KEY}`,
      'Content-Type': file.type || 'application/octet-stream',
      'x-upsert': 'true'
    },
    body: file
  });
  if (!res.ok) throw new Error(`Upload falhou (${res.status}): ${await res.text()}`);
  return { path, publicUrl: `${SUPABASE_URL}/storage/v1/object/public/${SGA_BUCKET}/${path}` };
};

const supaStorageDelete = async (storagePath) => {
  if (!storagePath) return;
  await fetch(`${SUPABASE_URL}/storage/v1/object/${SGA_BUCKET}/${storagePath}`, {
    method: 'DELETE',
    headers: { 'apikey': SUPABASE_ANON_KEY, 'Authorization': `Bearer ${SUPABASE_ANON_KEY}` }
  });
};

const formatFileSize = (bytes) => {
  if (!bytes) return '';
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
};

// ---------------------------------------------------------------------------
// CADASTRO DE USUÁRIOS — email → { nome, setor, role }
// ---------------------------------------------------------------------------
const SENHA_PADRAO = 'Kdb@2026';

const USUARIOS_BASE = [
  { email: 'luciene.batista@kalenborn.com.br', nome: 'Luciene Batista', setor: 'QUALIDADE - SEGURANÇA NO TRABALHO', role: 'admin' },
  { email: 'asael.abdon@kalenborn.com.br',    nome: 'Asael Abdon',     setor: 'QUALIDADE - SEGURANÇA NO TRABALHO', role: 'admin' },
  { email: 'luiz@kalenborn.com.br',            nome: 'Luiz',            setor: 'ALTA DIREÇÃO', role: 'setor' },
  { email: 'fabio.lago@kalenborn.com.br',       nome: 'Fábio Lago',     setor: 'ALTA DIREÇÃO', role: 'setor' },
  { email: 'ricardo.pereira@kalenborn.com.br',  nome: 'Ricardo Pereira',   setor: 'COMERCIAL', role: 'setor' },
  { email: 'priscila.monara@kalenborn.com.br',  nome: 'Priscila Monara',   setor: 'COMERCIAL', role: 'setor' },
  { email: 'jose.martins@kalenborn.com.br',     nome: 'José Martins',      setor: 'COMERCIAL', role: 'admin' },
  { email: 'leonardo.henrique@kalenborn.com.br', nome: 'Leonardo Henrique', setor: 'COMPRAS', role: 'admin' },
  { email: 'franciele.dias@kalenborn.com.br',    nome: 'Franciele Dias',    setor: 'COMPRAS', role: 'setor' },
  { email: 'danilo.lopes@kalenborn.com.br',      nome: 'Danilo Lopes',      setor: 'COMPRAS', role: 'setor' },
  { email: 'logistica@kalenborn.com.br',         nome: 'Logística',         setor: 'LOGÍSTICA', role: 'setor' },
  { email: 'welington.silva@kalenborn.com.br',   nome: 'Welington Silva',   setor: 'FISCAL', role: 'setor' },
  { email: 'sergio.malaquias@kalenborn.com.br',  nome: 'Sérgio Malaquias',  setor: 'INSPEÇÃO DE QUALIDADE', role: 'setor' },
  { email: 'daniel.toledo@kalenborn.com.br',     nome: 'Daniel Toledo',     setor: 'MANUTENÇÃO', role: 'admin' },
  { email: 'diogo.ribeiro@kalenborn.com.br',     nome: 'Diogo Ribeiro',     setor: 'PRODUÇÃO - CALDERARIA', role: 'setor' },
  { email: 'edson.menezes@kalenborn.com.br',     nome: 'Edson Menezes',     setor: 'ORÇAMENTO E DESENVOLVIMENTO', role: 'setor' },
  { email: 'melina.nunes@kalenborn.com.br',      nome: 'Melina Nunes',      setor: 'RECURSOS HUMANOS', role: 'setor' },
  { email: 'mariele.xavier@kalenborn.com.br',    nome: 'Mariele Xavier',    setor: 'RECURSOS HUMANOS', role: 'setor' },
];

const MULTI_SETOR = {
  'fabio.lago@kalenborn.com.br':    ['ALTA DIREÇÃO', 'FISCAL', 'RECURSOS HUMANOS'],
  'danilo.lopes@kalenborn.com.br':  ['COMPRAS', 'LOGÍSTICA'],
  'daniel.toledo@kalenborn.com.br': ['MANUTENÇÃO', 'PCP', 'PRODUÇÃO - CALDERARIA', 'PRODUÇÃO - REVESTIMENTO', 'PRODUÇÃO - VULCANIZAÇÃO'],
  'edson.menezes@kalenborn.com.br': ['ORÇAMENTO E DESENVOLVIMENTO', 'PROJETO - DESENHO'],
};

const findUsuario = (email) => {
  const e = email.trim().toLowerCase();
  return USUARIOS_BASE.find((u) => u.email.toLowerCase() === e) || null;
};

const getUserSetores = (usuario) => {
  if (!usuario) return [];
  if (usuario.role === 'admin') return ['TODOS'];
  return MULTI_SETOR[usuario.email.toLowerCase()] || [usuario.setor];
};

// ---------------------------------------------------------------------------
// COMPONENTE DE LOGIN
// ---------------------------------------------------------------------------
const LoginScreen = ({ onLogin, extraAdmins = [] }) => {
  const [email, setEmail] = React.useState('');
  const [senha, setSenha] = React.useState('');
  const [erro, setErro] = React.useState('');
  const [loading, setLoading] = React.useState(false);
  const [showPass, setShowPass] = React.useState(false);

  const handleLogin = async (e) => {
    e.preventDefault();
    setErro('');
    setLoading(true);
    await new Promise((r) => setTimeout(r, 400));

    if (senha !== SENHA_PADRAO) {
      setErro('Senha incorreta. Verifique e tente novamente.');
      setLoading(false);
      return;
    }
    const usuario = findUsuario(email);
    if (!usuario) {
      setErro('E-mail não cadastrado no sistema. Contate a Gestão da Qualidade.');
      setLoading(false);
      return;
    }
    const isExtraAdmin = extraAdmins.includes(usuario.email.toLowerCase());
    onLogin(isExtraAdmin ? { ...usuario, role: 'admin' } : usuario);
    setLoading(false);
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-[#0B1020] via-[#10142B] to-[#0B1020] flex items-center justify-center p-4">
      <div className="w-full max-w-md">
        <div className="flex justify-center mb-10">
          <img
            src={KALEN_LOGO_BRANCA_SRC}
            alt="Kalenborn Wear Protection Solutions"
            style={{ height: '60px', width: 'auto', objectFit: 'contain' }}
          />
        </div>
        <div className="bg-white/5 backdrop-blur border border-white/10 rounded-3xl p-8 shadow-2xl">
          <div className="mb-6">
            <h1 className="text-2xl font-black text-white tracking-tight">Portal da Qualidade</h1>
            <p className="text-slate-400 text-sm mt-1">Sistema de Gestão de Auditoria — Unidade Vespasiano</p>
          </div>
          <form onSubmit={handleLogin} className="space-y-4">
            <div>
              <label className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-1.5 block">E-mail corporativo</label>
              <input type="email" value={email} onChange={(e) => setEmail(e.target.value)} placeholder="seu.nome@kalenborn.com.br" required autoComplete="email"
                className="w-full bg-white/10 border border-white/10 text-white placeholder-slate-500 rounded-xl p-3 text-sm focus:ring-2 focus:ring-indigo-500 focus:outline-none" />
            </div>
            <div>
              <label className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-1.5 block">Senha</label>
              <div className="relative">
                <input type={showPass ? 'text' : 'password'} value={senha} onChange={(e) => setSenha(e.target.value)} placeholder="••••••••" required autoComplete="current-password"
                  className="w-full bg-white/10 border border-white/10 text-white placeholder-slate-500 rounded-xl p-3 text-sm focus:ring-2 focus:ring-indigo-500 focus:outline-none pr-20" />
                <button type="button" onClick={() => setShowPass((v) => !v)}
                  className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-white transition text-xs font-bold">
                  {showPass ? 'OCULTAR' : 'VER'}
                </button>
              </div>
            </div>
            {erro && (
              <div className="bg-rose-500/10 border border-rose-500/30 text-rose-400 rounded-xl p-3 text-sm flex items-start gap-2">
                <AlertCircle size={16} className="mt-0.5 shrink-0" />{erro}
              </div>
            )}
            <button type="submit" disabled={loading}
              className={`w-full py-3.5 rounded-xl font-bold text-sm transition flex items-center justify-center gap-2 ${loading ? 'bg-indigo-600/50 text-white/50 cursor-not-allowed' : 'bg-indigo-600 hover:bg-indigo-500 text-white shadow-lg shadow-indigo-600/30'}`}>
              {loading ? <><span className="animate-spin w-4 h-4 border-2 border-white/30 border-t-white rounded-full inline-block"></span> Entrando...</> : 'Entrar no Portal'}
            </button>
          </form>
          <p className="text-center text-slate-500 text-xs mt-6">
            Problemas de acesso? Contate <span className="text-slate-300 font-bold">luciene.batista@kalenborn.com.br</span>
          </p>
        </div>
        <p className="text-center text-slate-600 text-[10px] mt-4 uppercase tracking-wider">
          Kalenborn do Brasil · SGA v2.0 · ISO 9001:2015
        </p>
      </div>
    </div>
  );
};

// Fallback local enquanto dados do banco carregam (evita tela branca no boot)
const AUDITORES = ['Luciene Batista'];
const BRANCHES = ['Unidade Vespasiano - MG'];

// Lista de setores estática inicial — será substituída pelos dados do banco
const SECTORS_FALLBACK = [
  'ALTA DIREÇÃO', 'COMERCIAL', 'COMPRAS', 'FISCAL', 'INSPEÇÃO DE QUALIDADE',
  'LOGÍSTICA', 'MANUTENÇÃO', 'ORÇAMENTO E DESENVOLVIMENTO', 'PCP',
  'PRODUÇÃO - CALDERARIA', 'PRODUÇÃO - REVESTIMENTO', 'PRODUÇÃO - VULCANIZAÇÃO',
  'PROJETO - DESENHO', 'QUALIDADE - SEGURANÇA NO TRABALHO', 'RECURSOS HUMANOS'
];

// Checklist genérico — usado somente como fallback se o banco não responder
const GENERIC_CHECKLIST = [
  'Como você trabalha para atender esta política?',
  'Como é o funcionamento do setor? Em documento se orienta? (Verificar o fluxo do processo).',
  'Quais são as entradas para o processo?',
  'Quais são as saídas do processo?',
  'Como é gerenciado o processo no dia a dia?',
  'Quais são os registros do setor? (Verificar os registros).',
  'Verificar os indicadores. Caso um indicador esteja fora da meta, qual a tratativa?',
  'Os processos de retrabalho são direcionados para abertura de registro de não conformidade?'
];

// buildChecklist agora recebe o mapa dinâmico de templates vindo do banco
const buildChecklist = (sector, templateMap = {}) => {
  const questions = templateMap[sector] || GENERIC_CHECKLIST;
  return questions.map((question, idx) => ({
    id: idx + 1,
    question,
    status: '',
    comments: ''
  }));
};


let _checklistItemSeq = 10000; // ids para itens adicionados manualmente em tela
const nextChecklistItemId = () => ++_checklistItemSeq;

let _eventSeq = 5000; // ids para eventos do calendário de auditorias
const nextEventId = () => ++_eventSeq;

// ---------------------------------------------------------------------------
// UTILITÁRIOS
// ---------------------------------------------------------------------------

const pad3 = (n) => String(n).padStart(3, '0');
const pad2 = (n) => String(n).padStart(2, '0');

// ---- Utilitários de data para o calendário (sem libs externas) ----
const WEEKDAY_LABELS = ['Dom', 'Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb'];
const MONTH_LABELS = [
  'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
  'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'
];

const toDateKey = (date) => `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}`;

const dateKeyToDate = (key) => {
  const [y, m, d] = key.split('-').map(Number);
  return new Date(y, m - 1, d);
};

const addDays = (date, days) => {
  const d = new Date(date);
  d.setDate(d.getDate() + days);
  return d;
};

const startOfWeek = (date) => addDays(date, -date.getDay());

// Gera a grade de 6 semanas (42 dias) para a visão mensal, incluindo dias
// do mês anterior/seguinte para preencher as semanas completas.
const getMonthGrid = (year, month) => {
  const firstOfMonth = new Date(year, month, 1);
  const gridStart = startOfWeek(firstOfMonth);
  return Array.from({ length: 42 }, (_, i) => addDays(gridStart, i));
};

const getWeekGrid = (anchorDate) => {
  const start = startOfWeek(anchorDate);
  return Array.from({ length: 7 }, (_, i) => addDays(start, i));
};

const isSameDay = (a, b) => toDateKey(a) === toDateKey(b);

const formatTimeRange = (start, durationMin) => {
  const [h, m] = start.split(':').map(Number);
  const startMin = h * 60 + m;
  const endMin = startMin + durationMin;
  const fmt = (totalMin) => `${pad2(Math.floor(totalMin / 60) % 24)}:${pad2(totalMin % 60)}`;
  return `${fmt(startMin)} – ${fmt(endMin)}`;
};



// ---------------------------------------------------------------------------
// COMPONENTE PRINCIPAL
// ---------------------------------------------------------------------------

// Wrapper que controla login/logout — mantém o estado de auth fora do App
// para evitar violação das regras de Hooks (useState antes de return condicional)
export default function AppWrapper() {
  const [currentUser, setCurrentUser] = useState(null);
  // extraAdmins: emails promovidos a auditor pela Luciene — persiste no Supabase
  const [extraAdmins, setExtraAdmins] = useState([]);

  // Carrega extraAdmins do banco ao montar
  useEffect(() => {
    supaFetch('sga_extra_admins?select=email')
      .then((rows) => setExtraAdmins(rows.map((r) => r.email.toLowerCase())))
      .catch((err) => console.error('Erro ao carregar auditores promovidos:', err));
  }, []);

  // Promove/rebaixa um e-mail a admin — atualiza o banco e o estado local
  const toggleExtraAdmin = async (email, isPromovido) => {
    const e = email.toLowerCase();
    // Atualiza a UI imediatamente (otimista)
    setExtraAdmins((prev) => (isPromovido ? prev.filter((x) => x !== e) : [...prev, e]));
    try {
      if (isPromovido) {
        await supaFetch(`sga_extra_admins?email=eq.${encodeURIComponent(e)}`, { method: 'DELETE' });
      } else {
        await supaFetch('sga_extra_admins', { method: 'POST', body: JSON.stringify({ email: e }) });
      }
    } catch (err) {
      console.error('Erro ao atualizar auditor promovido no banco:', err);
      // Reverte a UI se o banco falhar
      setExtraAdmins((prev) => (isPromovido ? [...prev, e] : prev.filter((x) => x !== e)));
    }
  };

  const handleLogin = (usuario) => setCurrentUser(usuario);
  const handleLogout = () => setCurrentUser(null);

  if (!currentUser) {
    return (
      <ErrorBoundary>
        <LoginScreen onLogin={handleLogin} extraAdmins={extraAdmins} />
      </ErrorBoundary>
    );
  }

  return (
    <ErrorBoundary>
      <App
        currentUser={currentUser}
        extraAdmins={extraAdmins}
        toggleExtraAdmin={toggleExtraAdmin}
        onLogout={handleLogout}
      />
    </ErrorBoundary>
  );
}

function App({ currentUser, extraAdmins, toggleExtraAdmin, onLogout }) {
  // Bug 1 fix: isAdmin calculado aqui dentro, reativo a mudanças em extraAdmins
  const isAdmin = currentUser.role === 'admin' || extraAdmins.includes(currentUser.email.toLowerCase());
  const userSetores = getUserSetores({ ...currentUser, role: isAdmin ? 'admin' : currentUser.role });

  const [activeTab, setActiveTab] = useState(isAdmin ? 'minhas_auditorias' : 'documentos');
  const [showAdminPanel, setShowAdminPanel] = useState(false);
  const [selectedBranch, setSelectedBranch] = useState(BRANCHES[0]);

  // ---- Dados dinâmicos vindos do Supabase ----
  // sectors: [{ nome, requisitos }]  |  templateMap: { 'SETOR': ['pergunta1', ...] }
  const [sectors, setSectors] = useState(
    SECTORS_FALLBACK.map((nome) => ({ nome, requisitos: '' }))
  );
  const [templateMap, setTemplateMap] = useState({});
  // Lista real de auditores cadastrados no banco (sga_auditores) — usada para
  // vincular auditor_id ao salvar auditorias/eventos. Antes disso era uma
  // lista fixa no código com só a Luciene, então ninguém mais ficava vinculado.
  const [auditoresDb, setAuditoresDb] = useState([{ id: 1, nome: 'Luciene Batista' }]);
  const [dbLoading, setDbLoading] = useState(true);
  const [dbError, setDbError] = useState(null);
  // Estatísticas agregadas de TODAS as auditorias já concluídas (histórico completo)
  const [historicalStats, setHistoricalStats] = useState({ c: 0, nc: 0, obs: 0, naoAvaliado: 0, total: 0 });

  useEffect(() => {
    const loadFromSupabase = async () => {
      // Cada seção tem seu próprio try/catch: se uma falhar, as outras continuam
      // carregando normalmente em vez de deixar a tela inteira vazia.

      // 1. Setores com requisitos
      try {
        const setoresData = await supaFetch('sga_setores?select=id,nome,requisitos&ativo=eq.true&order=nome');
        if (setoresData.length > 0) setSectors(setoresData);
      } catch (err) {
        console.error('Erro ao carregar setores:', err);
        setDbError('Falha ao carregar setores.');
      }

      // 1b. Auditores cadastrados
      try {
        const auditoresData = await supaFetch('sga_auditores?select=id,nome&order=nome');
        if (auditoresData.length > 0) setAuditoresDb(auditoresData);
      } catch (err) {
        console.error('Erro ao carregar auditores:', err);
        setDbError('Falha ao carregar auditores.');
      }

      // 2. Templates de perguntas (todas de uma vez, ordenadas)
      try {
        const templatesData = await supaFetch(
          'sga_checklist_templates?select=pergunta,ordem,sga_setores(nome)&order=ordem'
        );
        const map = {};
        templatesData.forEach((t) => {
          const setor = t.sga_setores?.nome;
          if (!setor) return;
          if (!map[setor]) map[setor] = [];
          map[setor].push(t.pergunta);
        });
        setTemplateMap(map);
      } catch (err) {
        console.error('Erro ao carregar templates de checklist:', err);
        setDbError('Falha ao carregar templates de checklist.');
      }

      // 3. Auditorias salvas (histórico)
      try {
        const auditoriasData = await supaFetch(
          'sga_auditorias?select=id,rai_numero,data_emissao,unidade,status,qtd_nc,auditado_nome,pontos_positivos,observacoes,melhorias,conclusao,sga_setores(nome),sga_auditores(nome)&order=created_at.desc&limit=100'
        );
        if (auditoriasData.length > 0) {
          const mapped = auditoriasData.map((a) => ({
            id: a.id,
            dbId: a.id,
            raiNumber: a.rai_numero,
            date: a.data_emissao ? new Date(a.data_emissao).toLocaleDateString('pt-BR', { day:'2-digit', month:'2-digit', year:'2-digit' }) : '—',
            sector: a.sga_setores?.nome || '—',
            branch: a.unidade || BRANCHES[0],
            auditor: a.sga_auditores?.nome || '—',
            ncCount: a.qtd_nc || 0,
            status: a.status || 'Concluído',
            // Campos brutos do relatório — usados para reconstruir reportSnapshot no "VER"
            // sem precisar salvar o objeto inteiro no cliente (checklist ainda é buscado sob demanda)
            _dbReport: {
              raiNumber: a.rai_numero,
              date: a.data_emissao ? new Date(a.data_emissao).toLocaleDateString('pt-BR', { day:'2-digit', month:'2-digit', year:'2-digit' }) : '',
              auditor: a.sga_auditores?.nome || '',
              auditee: a.auditado_nome || '',
              positivePoints: a.pontos_positivos || '',
              observations: a.observacoes || '',
              improvements: a.melhorias || '',
              conclusion: a.conclusao || ''
            }
          }));
          setSavedAudits(mapped);
          // Ajusta o contador RAI ao maior número encontrado — update funcional
          // porque o cálculo dos rascunhos (outra seção) roda em paralelo e não
          // pode ser sobrescrito por este aqui (nem o contrário).
          const nums = mapped.map((a) => parseInt(a.raiNumber?.split('/')[0] || '0', 10)).filter(Boolean);
          if (nums.length) setRaiCounter(prev => Math.max(prev, Math.max(...nums) + 1));
        }
      } catch (err) {
        console.error('Erro ao carregar histórico de auditorias:', err);
        setDbError('Falha ao carregar histórico de auditorias.');
      }

      // 4. RNCs
      try {
        const rncsData = await supaFetch(
          'sga_rncs?select=*,sga_auditorias(rai_numero),sga_checklist_itens(id,ordem,pergunta)&order=created_at.desc&limit=200'
        );
        if (rncsData.length > 0) {
          const mappedRncs = rncsData.map((r) => ({
            id: r.rac_numero || String(r.id),
            dbId: r.id,
            date: r.data_emissao ? new Date(r.data_emissao).toLocaleDateString('pt-BR', { day:'2-digit', month:'2-digit', year:'2-digit' }) : '—',
            process: r.processo || '',
            origin: r.origem || 'AUDITORIA INTERNA',
            // Reconstituídos a partir dos vínculos do banco (auditoria_id / checklist_item_id)
            // — sem isso, depois de recarregar a página o app "esquecia" que um item de
            // checklist já tinha RNC aberta e deixava criar duplicada.
            sourceRaiNumber: r.sga_auditorias?.rai_numero || null,
            sourceChecklistItem: r.sga_checklist_itens
              ? { id: r.sga_checklist_itens.id, question: r.sga_checklist_itens.pergunta }
              : null,
            description: r.descricao_nc || '',
            correction: r.correcao || '',
            correctionResp: r.correcao_resp || '',
            correctionDate: r.correcao_data_limite || '',
            rootCause: r.causa_raiz || '',
            rootCauseResp: r.causa_raiz_resp || '',
            rootCauseDate: r.causa_raiz_data || '',
            actionPlan: r.plano_acao || '',
            responsible: r.plano_responsavel || '',
            actionPlanDatePrev: r.plano_data_prevista || '',
            actionPlanDateReal: r.plano_data_realizada || '',
            evidence: r.evidencias || '',
            effective: r.eficaz || '',
            newRnc: r.necessita_nova_rnc || '',
            closeDate: r.data_fechamento || '',
            closeResp: r.validador_sgq || '',
            status: r.status || 'Aberta',
          }));
          setRncs(mappedRncs);
          // Ajusta o contador de RNCs
          const rncNums = mappedRncs.map((r) => parseInt(r.id, 10)).filter(Boolean);
          if (rncNums.length) setRncCounter(Math.max(...rncNums) + 1);
        }
      } catch (err) {
        console.error('Erro ao carregar RNCs:', err);
        setDbError('Falha ao carregar RNCs.');
      }

      // 5. Eventos do calendário
      try {
        const eventosData = await supaFetch(
          'sga_eventos_calendario?select=id,data_auditoria,horario,duracao_min,unidade,observacoes,status,sga_setores(nome),sga_auditores(nome)&order=data_auditoria&limit=100'
        );
        if (eventosData.length > 0) {
          const mappedEvents = eventosData.map((e) => ({
            id: e.id,
            dateKey: e.data_auditoria,
            time: e.horario?.slice(0, 5) || '09:00',
            durationMin: e.duracao_min || 60,
            sector: e.sga_setores?.nome || '',
            auditor: e.sga_auditores?.nome || AUDITORES[0],
            branch: e.unidade || BRANCHES[0],
            notes: e.observacoes || '',
          }));
          setAuditEvents(mappedEvents);
        }
      } catch (err) {
        console.error('Erro ao carregar eventos do calendário:', err);
        setDbError('Falha ao carregar calendário.');
      }

      // 6. Documentos por setor
      try {
        const docsData = await supaFetch(
          'sga_documentos?select=*&order=created_at.desc&limit=500'
        );
        if (docsData.length > 0) {
          const mappedDocs = docsData.map((d) => ({
            id: d.id,
            dbId: d.id,
            setor: d.setor,
            nome: d.nome,
            tipo: d.tipo,
            responsavel: d.responsavel || '',
            validade: d.validade || '',
            descricao: d.descricao || '',
            url: d.storage_path
              ? `${SUPABASE_URL}/storage/v1/object/public/${SGA_BUCKET}/${d.storage_path}`
              : (d.url || ''),
            storage_path: d.storage_path || null,
            storage_name: d.storage_name || null,
            storage_size: d.storage_size || null,
            storage_mime: d.storage_mime || null,
            created_at: d.created_at
          }));
          setDocumentos(mappedDocs);
        }
      } catch (err) {
        console.error('Erro ao carregar documentos:', err);
        setDbError('Falha ao carregar documentos.');
      }

      // 7. Estatísticas históricas agregadas — C/NC/Obs de TODAS as auditorias já
      // concluídas (sga_checklist_itens só tem itens de auditorias finalizadas,
      // rascunhos não entram aqui). Pagina automaticamente para pegar tudo,
      // não só as primeiras 1000 linhas.
      try {
        const itensAvaliacao = await supaFetchAll('sga_checklist_itens?select=avaliacao');
        const c   = itensAvaliacao.filter((i) => i.avaliacao === 'C').length;
        const nc  = itensAvaliacao.filter((i) => i.avaliacao === 'NC').length;
        const obs = itensAvaliacao.filter((i) => i.avaliacao === 'Obs').length;
        const naoAvaliado = itensAvaliacao.filter((i) => !i.avaliacao).length;
        setHistoricalStats({ c, nc, obs, naoAvaliado, total: itensAvaliacao.length });
      } catch (err) {
        console.error('Erro ao carregar estatísticas históricas:', err);
        setDbError('Falha ao carregar estatísticas do histórico.');
      }

      setDbLoading(false);
    };
    loadFromSupabase();
  }, []);

  // Atualiza o checklist quando os templates carregam do banco (se o setor já está selecionado)
  useEffect(() => {
    if (!dbLoading && Object.keys(templateMap).length > 0) {
      setChecklist(buildChecklist(selectedSector, templateMap));
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [dbLoading, templateMap]);

  // Lista de nomes de setores para os dropdowns
  const SECTORS = sectors.map((s) => s.nome);

  const [raiCounter, setRaiCounter] = useState(1);
  const [rncCounter, setRncCounter] = useState(1);
  const currentYear = new Date().getFullYear();
  const today = new Date();
  const formattedDate = today.toLocaleDateString('pt-BR', { day: '2-digit', month: '2-digit', year: '2-digit' });

  // ── ARQUITETURA: auditorias independentes por setor ───────────────────────
  // Cada auditoria tem setor fixo. Rascunhos ficam 100% no Supabase — nada em
  // localStorage. Cada salvamento (auto ou manual) INSERE uma versão nova
  // (tabela sga_rascunhos); nunca sobrescreve uma existente. O app sempre lê a
  // versão mais recente através da view sga_rascunhos_atual.
  //
  // IMPORTANTE sobre o desenho do estado: o autosave e o salvamento manual leem
  // os valores DIRETO dos states do editor (checklist, report, etc.) — nunca de
  // volta através do array `auditorias`. Ler de volta do array causava bug de
  // closure desatualizada (o save sempre gravava a versão anterior à edição
  // mais recente, porque os dois useEffects de mesma dependência rodavam com o
  // mesmo array "auditorias" ainda não atualizado). A identidade no banco
  // (dbId/versão) de quem está sendo editado agora vive numa ref, não no array.

  // 1. Lista de rascunhos — só para exibição em "Minhas Auditorias". Carregada
  // do banco ao montar; atualizada quando cria/remove/finaliza um rascunho.
  const [auditorias, setAuditorias] = useState([]);
  const [rascunhosLoaded, setRascunhosLoaded] = useState(false);

  const recarregarRascunhos = () => {
    return supaFetch(`sga_rascunhos_atual?usuario_email=eq.${encodeURIComponent(currentUser.email)}&select=*`)
      .then((rows) => {
        setAuditorias(rows.map((r) => ({
          localId: r.local_id,
          dbId: r.id,
          versao: r.versao,
          setor: r.setor,
          raiNumber: r.rai_numero,
          checklist: r.checklist || [],
          checklistStatus: r.checklist_status || 'Em Andamento',
          checklistClosedAt: r.checklist_closed_at,
          reopenHistory: r.reopen_history || [],
          report: r.report || {}
        })));
      })
      .catch((err) => console.error('Erro ao carregar rascunhos do banco:', err));
  };

  useEffect(() => {
    recarregarRascunhos().finally(() => setRascunhosLoaded(true));

    // Considera também os rascunhos de OUTROS auditores (não só os do usuário
    // atual) ao calcular o próximo número de RAI — evita que dois auditores
    // diferentes, cada um com um rascunho aberto, acabem reservando o mesmo
    // número de RAI sem saber um do outro.
    supaFetch('sga_rascunhos_atual?select=rai_numero')
      .then((rows) => {
        const nums = rows.map((r) => parseInt(r.rai_numero?.split('/')[0] || '0', 10)).filter(Boolean);
        if (nums.length) setRaiCounter(prev => Math.max(prev, Math.max(...nums) + 1));
      })
      .catch((err) => console.error('Erro ao calcular próximo número de RAI a partir dos rascunhos:', err));
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  // 2. ID da auditoria aberta (null = nenhuma). Só é usado pra rascunhos em
  // andamento — ao abrir uma auditoria já FINALIZADA (VER / editar histórico),
  // isso é sempre limpo para null, pra nunca misturar os dois fluxos.
  const [auditoriaAtivaId, setAuditoriaAtivaId] = useState(null);

  // Identidade da linha no banco do rascunho ativo — vive numa ref (não em
  // state) porque é lida de forma síncrona dentro dos callbacks de save, sem
  // precisar esperar um re-render.
  const auditoriaAtivaMetaRef = useRef({ dbId: null, versao: 1 });

  // 3. Estados independentes do editor — iniciam vazios, sincronizam com a ativa
  const [checklist,         setChecklist        ] = useState([]);
  const [selectedSector,    setSelectedSector   ] = useState('COMERCIAL');

  // Requisitos do setor selecionado (mostrado no cabeçalho do checklist e RAI)
  const selectedSectorRequisitos = sectors.find((s) => s.nome === selectedSector)?.requisitos || '';

  const [checklistStatus,   setChecklistStatus  ] = useState('Em Andamento');
  const [checklistClosedAt, setChecklistClosedAt] = useState(null);
  const [reopenHistory,     setReopenHistory    ] = useState([]);
  const [report,            setReport           ] = useState({});
  const [dynamicRaiNumber,  setDynamicRaiNumber ] = useState(`${pad3(1)}/${new Date().getFullYear()}`);

  // 4. Quando a auditoria ativa muda, sincroniza os estados do editor
  useEffect(() => {
    const ativa = auditorias.find(a => a.localId === auditoriaAtivaId) || null;
    if (ativa) {
      setChecklist(ativa.checklist || []);
      setSelectedSector(ativa.setor || 'COMERCIAL');
      setChecklistStatus(ativa.checklistStatus || 'Em Andamento');
      setChecklistClosedAt(ativa.checklistClosedAt ? new Date(ativa.checklistClosedAt) : null);
      setReopenHistory(ativa.reopenHistory || []);
      setReport(ativa.report || {});
      setDynamicRaiNumber(ativa.raiNumber || `${pad3(raiCounter)}/${currentYear}`);
      auditoriaAtivaMetaRef.current = { dbId: ativa.dbId || null, versao: ativa.versao || 1 };
    } else {
      setChecklist([]);
      setSelectedSector('COMERCIAL');
      setChecklistStatus('Em Andamento');
      setChecklistClosedAt(null);
      setReopenHistory([]);
      setReport({});
    }
  }, [auditoriaAtivaId]); // eslint-disable-line react-hooks/exhaustive-deps

  // 5. Quando os estados do editor mudam, atualiza a cópia exibida em "Minhas
  // Auditorias" (progresso, etc.). Isso é só para a LISTAGEM — o autosave (item
  // 6 abaixo) NÃO depende deste efeito, lê os states direto.
  useEffect(() => {
    if (!auditoriaAtivaId) return;
    setAuditorias(prev => prev.map(a =>
      a.localId !== auditoriaAtivaId ? a : {
        ...a,
        checklist,
        checklistStatus,
        checklistClosedAt: checklistClosedAt ? checklistClosedAt.toISOString() : null,
        reopenHistory,
        report
      }
    ));
  }, [checklist, checklistStatus, checklistClosedAt, reopenHistory, report]); // eslint-disable-line react-hooks/exhaustive-deps

  // Insere uma nova versão do rascunho ativo no banco (nunca sobrescreve a
  // anterior). Monta o payload DIRETO dos states do editor — nunca do array
  // `auditorias` — para nunca gravar uma versão desatualizada.
  const inserirVersaoRascunhoAtivo = async () => {
    const proximaVersao = (auditoriaAtivaMetaRef.current.versao || 1) + 1;
    const [savedRow] = await supaFetch('sga_rascunhos', {
      method: 'POST',
      body: JSON.stringify({
        local_id: auditoriaAtivaId,
        usuario_email: currentUser.email,
        setor: selectedSector,
        rai_numero: dynamicRaiNumber,
        versao: proximaVersao,
        checklist,
        checklist_status: checklistStatus,
        checklist_closed_at: checklistClosedAt ? checklistClosedAt.toISOString() : null,
        reopen_history: reopenHistory,
        report
      })
    });
    if (savedRow?.id) {
      auditoriaAtivaMetaRef.current = { dbId: savedRow.id, versao: proximaVersao };
      setAuditorias(prev => prev.map(a =>
        a.localId === auditoriaAtivaId ? { ...a, dbId: savedRow.id, versao: proximaVersao } : a
      ));
    }
  };

  // 6. Autosave no Supabase — debounced (1.5s de inatividade) para não gerar
  // uma requisição por tecla digitada. Cada save cria uma versão nova (insert),
  // nunca sobrescreve a versão anterior. Só roda depois do carregamento inicial.
  const rascunhoSaveTimer = useRef(null);
  useEffect(() => {
    if (!auditoriaAtivaId || !rascunhosLoaded) return;

    setAutoSaveStatus('unsaved');
    clearTimeout(rascunhoSaveTimer.current);
    rascunhoSaveTimer.current = setTimeout(async () => {
      setAutoSaveStatus('saving');
      try {
        await inserirVersaoRascunhoAtivo();
        setAutoSaveStatus('saved');
        setLastSavedAt(new Date());
      } catch (err) {
        console.error('Erro ao salvar nova versão do rascunho no banco:', err);
        setAutoSaveStatus('error');
      }
    }, 1500);
    return () => clearTimeout(rascunhoSaveTimer.current);
  }, [checklist, checklistStatus, checklistClosedAt, reopenHistory, report]); // eslint-disable-line react-hooks/exhaustive-deps

  // ── Indicador de status ───────────────────────────────────────────────────
  const [autoSaveStatus, setAutoSaveStatus] = useState('saved');
  const [lastSavedAt, setLastSavedAt]       = useState(null);
  const markUnsaved = () => setAutoSaveStatus('unsaved');

  // Salvamento manual imediato (botão "Salvar Rascunho") — ignora o debounce,
  // mas segue a mesma regra: insere uma versão nova, nunca sobrescreve.
  const saveRascunho = async () => {
    if (!auditoriaAtivaId) return;
    clearTimeout(rascunhoSaveTimer.current);
    setAutoSaveStatus('saving');
    try {
      await inserirVersaoRascunhoAtivo();
      setAutoSaveStatus('saved');
      setLastSavedAt(new Date());
    } catch (err) {
      console.error('Erro ao salvar nova versão do rascunho no banco:', err);
      setAutoSaveStatus('error');
      notify('⚠️ Erro ao salvar rascunho', `Não foi possível salvar no banco. Erro: ${err.message}`);
    }
  };

  // ── Criar nova auditoria ─────────────────────────────────────────────────
  const criarNovaAuditoria = async (setor) => {
    const localId = `aud_${Date.now()}`;
    const raiNum  = `${pad3(raiCounter)}/${currentYear}`;
    const nova = {
      localId,
      setor,
      raiNumber: raiNum,
      versao: 1,
      checklist: buildChecklist(setor, templateMap),
      checklistStatus: 'Em Andamento',
      checklistClosedAt: null,
      reopenHistory: [],
      report: {
        raiNumber: raiNum,
        date: today.toISOString().split('T')[0],
        auditor: currentUser.nome || AUDITORES[0],
        auditee: '', positivePoints: '', observations: '', improvements: '', conclusion: ''
      }
    };
    setRaiCounter(prev => prev + 1);
    setAuditorias(prev => [...prev, nova]);
    auditoriaAtivaMetaRef.current = { dbId: null, versao: 1 };
    setAuditoriaAtivaId(localId);
    setActiveTab('checklist');

    // Cria a versão 1 no banco imediatamente — antes mesmo do usuário digitar algo
    try {
      const [savedRow] = await supaFetch('sga_rascunhos', {
        method: 'POST',
        body: JSON.stringify({
          local_id: localId,
          usuario_email: currentUser.email,
          setor: nova.setor,
          rai_numero: nova.raiNumber,
          versao: 1,
          checklist: nova.checklist,
          checklist_status: nova.checklistStatus,
          reopen_history: nova.reopenHistory,
          report: nova.report
        })
      });
      if (savedRow?.id) {
        auditoriaAtivaMetaRef.current = { dbId: savedRow.id, versao: 1 };
        setAuditorias(prev => prev.map(a => a.localId === localId ? { ...a, dbId: savedRow.id } : a));
      }
    } catch (err) {
      console.error('Erro ao criar rascunho no banco:', err);
      notify('⚠️ Rascunho não salvo no banco', `A auditoria foi criada apenas nesta tela. Erro: ${err.message}. Tente recarregar a página.`);
    }
  };

  // ── Remover rascunho ─────────────────────────────────────────────────────
  // Apaga TODAS as versões desse rascunho no banco (é um descarte intencional,
  // diferente do autosave — aqui não faz sentido manter histórico).
  const removerRascunho = async (localId) => {
    setAuditorias(prev => prev.filter(a => a.localId !== localId));
    if (auditoriaAtivaId === localId) {
      setAuditoriaAtivaId(null);
      setActiveTab('minhas_auditorias');
    }
    try {
      await supaFetch(`sga_rascunhos?local_id=eq.${encodeURIComponent(localId)}`, { method: 'DELETE' });
    } catch (err) {
      console.error('Erro ao excluir rascunho do banco:', err);
    }
  };

  const [rncs, setRncs] = useState([]);
  const [savedAudits, setSavedAudits] = useState([]);
  const [loadingAuditId, setLoadingAuditId] = useState(null);
  const [historySearch, setHistorySearch] = useState('');
  const [historyPage, setHistoryPage]     = useState(1);
  const HISTORY_PAGE_SIZE = 8;
  const [rncStatusFilter, setRncStatusFilter] = useState('Todas');
  const [historyNcFilter, setHistoryNcFilter] = useState('Todas');
  const [editingAuditId, setEditingAuditId]   = useState(null);

  const [calendarView, setCalendarView] = useState('month');
  const [calendarAnchor, setCalendarAnchor] = useState(new Date());
  const [auditEvents, setAuditEvents] = useState([]);

  const [eventDraft, setEventDraft] = useState(null);
  const [draggedEventId, setDraggedEventId] = useState(null);
  const [dragOverDateKey, setDragOverDateKey] = useState(null);

  const [emailDraft, setEmailDraft] = useState(null);
  const [emailCopied, setEmailCopied] = useState(false);

  // ---- Documentos por setor ----
  const [documentos, setDocumentos] = useState([]);
  const [docDraft, setDocDraft] = useState(null);       // documento sendo criado/editado no modal
  const [docUploadFile, setDocUploadFile] = useState(null);  // File selecionado para upload
  const [docUploading, setDocUploading] = useState(false);   // progresso de upload
  const [docUploadError, setDocUploadError] = useState('');
  const [docSectorFilter, setDocSectorFilter] = useState(isAdmin ? 'TODOS' : (userSetores[0] || 'TODOS'));
  const [docSearch, setDocSearch] = useState('');

  const DOC_TIPOS = [
    'Procedimento', 'Instrução de Trabalho', 'Formulário', 'Registro',
    'Norma / Requisito', 'Manual', 'Plano de Ação', 'Certificado', 'Outro'
  ];

  const getDocStatus = (validade) => {
    if (!validade) return null;
    const hoje = new Date();
    const val = new Date(validade + 'T00:00:00');
    const diff = Math.ceil((val - hoje) / (1000 * 60 * 60 * 24));
    if (diff < 0) return 'vencido';
    if (diff <= 30) return 'proximo';
    return 'ok';
  };

  const filteredDocumentos = useMemo(() => {
    return documentos.filter((d) => {
      // Restrição de acesso: usuário de setor só vê seus setores
      const matchPermissao = isAdmin || userSetores.includes(d.setor);
      // Filtro de setor selecionado (dropdown/chip)
      const matchSetor = docSectorFilter === 'TODOS' || d.setor === docSectorFilter;
      const term = docSearch.trim().toLowerCase();
      const matchSearch = !term ||
        d.nome.toLowerCase().includes(term) ||
        d.setor.toLowerCase().includes(term) ||
        d.tipo.toLowerCase().includes(term) ||
        (d.responsavel || '').toLowerCase().includes(term);
      return matchPermissao && matchSetor && matchSearch;
    });
  }, [documentos, docSectorFilter, docSearch, isAdmin, userSetores]);

  // Carrega documentos do banco na inicialização (já acontece dentro do useEffect principal)

  // Salva documento: faz upload do arquivo se houver, depois persiste metadados no banco
  const saveDocumento = async (doc) => {
    setDocUploading(true);
    setDocUploadError('');
    try {
      let storagePath = doc.storage_path || null;
      let storagePublicUrl = doc.url || null;
      let storageName = doc.storage_name || null;
      let storageSize = doc.storage_size || null;
      let storageMime = doc.storage_mime || null;

      // 1. Upload do arquivo se um novo foi selecionado
      if (docUploadFile) {
        const { path, publicUrl } = await supaStorageUpload(docUploadFile);
        // Remove arquivo antigo do Storage se estava trocando
        if (doc.storage_path && doc.storage_path !== path) {
          await supaStorageDelete(doc.storage_path).catch(() => {});
        }
        storagePath = path;
        storagePublicUrl = publicUrl;
        storageName = docUploadFile.name;
        storageSize = docUploadFile.size;
        storageMime = docUploadFile.type;
      }

      const setorDocRow = sectors.find((s) => s.nome === doc.setor);
      const payload = {
        setor_id: setorDocRow?.id || null,
        nome: doc.nome,
        tipo: doc.tipo,
        responsavel: doc.responsavel || null,
        validade: doc.validade || null,
        descricao: doc.descricao || null,
        url_publica: storagePublicUrl || doc.url || null,
        storage_path: storagePath,
        storage_name: storageName,
        storage_size: storageSize,
        storage_mime: storageMime
      };

      const isNew = !doc.dbId;
      let savedDoc;
      if (isNew) {
        [savedDoc] = await supaFetch('sga_documentos', { method: 'POST', body: JSON.stringify(payload) });
      } else {
        [savedDoc] = await supaFetch(`sga_documentos?id=eq.${doc.dbId}`, { method: 'PATCH', body: JSON.stringify(payload) });
      }

      const finalDoc = {
        ...doc,
        ...payload,
        dbId: savedDoc?.id || doc.dbId,
        id: savedDoc?.id || doc.dbId || doc.id,
        // URL de acesso: arquivo do Storage tem prioridade sobre link externo
        url: storagePublicUrl || doc.url || null,
        url_publica: storagePublicUrl || doc.url || null,
        storage_path: storagePath,
        storage_name: storageName,
        storage_size: storageSize,
        storage_mime: storageMime,
        created_at: savedDoc?.created_at || doc.created_at || new Date().toISOString()
      };

      setDocumentos((prev) =>
        isNew ? [finalDoc, ...prev] : prev.map((d) => d.dbId === finalDoc.dbId ? finalDoc : d)
      );
      setDocDraft(null);
      setDocUploadFile(null);
    } catch (err) {
      console.error('Erro ao salvar documento:', err);
      setDocUploadError(err.message || 'Erro ao salvar documento. Verifique a conexão.');
    } finally {
      setDocUploading(false);
    }
  };

  const deleteDocumento = async (doc) => {
    // Remove da UI imediatamente
    setDocumentos((prev) => prev.filter((d) => d.dbId !== doc.dbId && d.id !== doc.id));
    // Remove arquivo do Storage (em background)
    if (doc.storage_path) supaStorageDelete(doc.storage_path).catch(() => {});
    // Remove registro do banco
    if (doc.dbId) {
      supaFetch(`sga_documentos?id=eq.${doc.dbId}`, { method: 'DELETE' }).catch(() => {});
    }
  };

  // Contadores para o badge da sidebar e alertas do dashboard
  const docsVencidos = documentos.filter((d) => getDocStatus(d.validade) === 'vencido').length;
  const docsProximos = documentos.filter((d) => getDocStatus(d.validade) === 'proximo').length;

  const [confirmAction, setConfirmAction] = useState(null);

  // ── Derivadas — declaradas DEPOIS de todos os useState ───────────────────
  // Isso evita TDZ na minificação (Vite/Rollup reorganiza closures)
  const auditoriaAtiva     = auditorias.find(a => a.localId === auditoriaAtivaId) || null;
  const usingGenericChecklist = !templateMap[selectedSector];

  const requestConfirm = (config) => setConfirmAction(config);
  const closeConfirm = () => setConfirmAction(null);
  const handleConfirmedAction = () => {
    if (confirmAction && typeof confirmAction.onConfirm === 'function') {
      confirmAction.onConfirm();
    }
    closeConfirm();
  };

  // Aviso simples (sem ação destrutiva) — reaproveita o mesmo modal
  const notify = (title, message) => requestConfirm({ title, message, confirmLabel: 'Entendi', danger: false, onConfirm: () => {} });

  // ---- Troca de setor: reconstrói o checklist e avisa se está usando o genérico ----
  // Setor é imutável após criar auditoria.
  // handleSectorChange agora cria nova auditoria para o setor escolhido.
  const handleSectorChange = (sector) => {
    if (!sector) return;
    criarNovaAuditoria(sector);
  };

  // Itens do checklist não podem ser editados enquanto estiver Fechado.
  const isChecklistLocked = checklistStatus === 'Fechado';

  const handleAddChecklistItem = () => {
    if (isChecklistLocked) return;
    setChecklist((prev) => [...prev, { id: nextChecklistItemId(), question: '', status: '', comments: '' }]);
  };

  const handleRemoveChecklistItem = (id) => {
    if (isChecklistLocked) return;
    setChecklist((prev) => prev.filter((item) => item.id !== id));
  };

  const closeChecklist = () => {
    setChecklistStatus('Fechado');
    setChecklistClosedAt(new Date());
  };

  const requestCloseChecklist = () => {
    const blockers = getSaveBlockers();
    if (blockers.length > 0) {
      notify('Não é possível fechar', blockers.join(' '));
      return;
    }
    requestConfirm({
      title: 'Fechar Lista de Verificação',
      message: `Tem certeza que deseja fechar a Lista de Verificação do setor ${selectedSector}? Após fechada, os itens ficam somente leitura até uma reabertura.`,
      confirmLabel: 'Sim, fechar checklist',
      danger: false,
      onConfirm: closeChecklist
    });
  };

  const requestReopenChecklist = () => {
    requestConfirm({
      title: 'Reabrir Lista de Verificação',
      message: 'Tem certeza que deseja reabrir esta Lista de Verificação para edição? A reabertura ficará registrada para rastreabilidade.',
      confirmLabel: 'Sim, reabrir checklist',
      danger: true,
      onConfirm: () => {
        setReopenHistory((prev) => [...prev, { date: new Date(), auditor: report.auditor }]);
        setChecklistStatus('Em Andamento');
        setChecklistClosedAt(null);
      }
    });
  };

  // ---- Validação de conclusão de auditoria ----
  const getSaveBlockers = () => {
    const blockers = [];
    if (getAnsweredCount() === 0) {
      blockers.push('Nenhum item do checklist foi avaliado ainda.');
    }
    const emptyQuestions = checklist.filter((i) => !i.question.trim());
    if (emptyQuestions.length > 0) {
      blockers.push(`${emptyQuestions.length} item(ns) do checklist está(ão) sem texto de pergunta.`);
    }
    const ncWithoutComment = checklist.filter((i) => i.status === 'NC' && !i.comments.trim());
    if (ncWithoutComment.length > 0) {
      blockers.push(`${ncWithoutComment.length} item(ns) marcado(s) como NC sem evidência/comentário registrado.`);
    }
    if (!report.auditor) {
      blockers.push('Selecione o auditor líder responsável.');
    }
    return blockers;
  };

  const handleSaveAudit = async () => {
    const blockers = getSaveBlockers();
    if (blockers.length > 0) {
      notify('Não foi possível salvar', blockers.join(' '));
      return;
    }

    if (checklistStatus !== 'Fechado') closeChecklist();

    const isEditing = !!editingAuditId;
    const localId = isEditing ? editingAuditId : Date.now();

    // Captura snapshot dos dados ANTES de qualquer alteração de estado
    const checklistSnapshot = JSON.parse(JSON.stringify(checklist));
    const reportSnapshot    = JSON.parse(JSON.stringify(report));
    const sectorSnapshot    = selectedSector;
    const branchSnapshot    = selectedBranch;
    const ncCountSnapshot   = getNcCount();
    const raiSnapshot       = report.raiNumber;

    // ── PASSO 1: Salva no Supabase ──────────────────────────────────────
    // Só avança se o banco confirmar. Se falhar, mostra erro e NÃO limpa nada.
    let dbId = isEditing ? editingAuditId : null;
    try {
      const setorRow = sectors.find((s) => s.nome === sectorSnapshot);
      const auditorRow = auditoresDb.find((a) => a.nome === reportSnapshot.auditor);
      const auditoriaPayload = {
        rai_numero:       raiSnapshot,
        setor_id:         setorRow?.id || null,
        auditor_id:       auditorRow?.id || null,
        unidade:          branchSnapshot,
        auditado_nome:    reportSnapshot.auditee,
        data_emissao:     reportSnapshot.date,
        status_checklist: 'Fechado',
        pontos_positivos: reportSnapshot.positivePoints,
        observacoes:      reportSnapshot.observations,
        melhorias:        reportSnapshot.improvements,
        conclusao:        reportSnapshot.conclusion,
        qtd_nc:           ncCountSnapshot,
        status:           'Concluído'
      };

      if (isEditing && typeof editingAuditId === 'number') {
        await supaFetch(`sga_auditorias?id=eq.${editingAuditId}`, {
          method: 'PATCH',
          body: JSON.stringify(auditoriaPayload)
        });
        await supaFetch(`sga_checklist_itens?auditoria_id=eq.${editingAuditId}`, { method: 'DELETE' });
      } else {
        const [savedAuditDb] = await supaFetch('sga_auditorias', {
          method: 'POST',
          body: JSON.stringify(auditoriaPayload)
        });
        if (savedAuditDb?.id) {
          dbId = savedAuditDb.id;
        } else {
          throw new Error('Banco não retornou ID da auditoria. Verifique as permissões do Supabase.');
        }
      }

      if (dbId) {
        const itensPayload = checklistSnapshot.map((item, idx) => ({
          auditoria_id: dbId,
          ordem:        idx + 1,
          pergunta:     item.question,
          avaliacao:    item.status || '',
          comentarios:  item.comments
        }));
        const insertedItens = await supaFetch('sga_checklist_itens', {
          method: 'POST',
          body: JSON.stringify(itensPayload)
        });

        // Vincula RNCs criadas durante esta auditoria ao ID real do item de
        // checklist no banco (checklist_item_id) — só existe a partir de agora,
        // porque os itens só ganham ID definitivo quando a auditoria fecha.
        // Sem isso, a rastreabilidade RNC → item do checklist se perdia após
        // recarregar a página (o vínculo só existia na memória da sessão).
        if (insertedItens?.length === itensPayload.length) {
          const idMap = {}; // id local do item (na tela) -> id real no banco
          checklistSnapshot.forEach((item, idx) => {
            idMap[item.id] = insertedItens[idx]?.id;
          });
          const rncsParaVincular = rncs.filter((r) =>
            r.sourceChecklistItem &&
            (r.sourceAuditoriaLocalId === auditoriaAtivaId || (isEditing && r.sourceAuditoriaLocalId === editingAuditId)) &&
            idMap[r.sourceChecklistItem.id] &&
            r.dbId
          );
          await Promise.all(rncsParaVincular.map((r) =>
            supaFetch(`sga_rncs?id=eq.${r.dbId}`, {
              method: 'PATCH',
              body: JSON.stringify({ checklist_item_id: idMap[r.sourceChecklistItem.id] })
            }).catch((e) => console.error(`Erro ao vincular RNC ${r.id} ao item do checklist:`, e))
          ));
        }
      }

    } catch (err) {
      // Banco falhou — NÃO limpa nada, NÃO reseta o checklist
      console.error('Erro ao persistir auditoria no Supabase:', err);
      notify(
        '⚠️ Erro ao salvar no banco',
        `Os dados NÃO foram salvos no banco. Não feche o navegador. Erro: ${err.message}. Tente novamente ou contate o suporte.`
      );
      return; // ← sai da função sem limpar nada
    }

    // ── PASSO 2: Banco confirmado — agora atualiza a UI e limpa o rascunho ──
    const auditData = {
      id:               dbId || localId,
      dbId:             dbId,
      isLocal:          false,
      raiNumber:        raiSnapshot,
      date:             formattedDate,
      sector:           sectorSnapshot,
      branch:           branchSnapshot,
      auditor:          reportSnapshot.auditor,
      ncCount:          ncCountSnapshot,
      status:           'Concluído',
      checklistSnapshot,
      reportSnapshot
    };

    if (isEditing) {
      setSavedAudits((prev) => prev.map((a) =>
        (a.dbId === editingAuditId || a.id === editingAuditId) ? { ...a, ...auditData } : a
      ));
    } else {
      setSavedAudits((prev) => [auditData, ...prev]);
    }

    setEditingAuditId(null);
    setAutoSaveStatus('saved');
    setLastSavedAt(new Date());

    // Remove auditoria da lista de rascunhos em andamento (foi concluída)
    if (!isEditing && auditoriaAtivaId) {
      setAuditorias(prev => prev.filter(a => a.localId !== auditoriaAtivaId));
      // Apaga TODAS as versões do rascunho no banco — a auditoria já está definitivamente
      // salva em sga_auditorias, não faz sentido manter o histórico de rascunho
      try {
        await supaFetch(`sga_rascunhos?local_id=eq.${encodeURIComponent(auditoriaAtivaId)}`, { method: 'DELETE' });
      } catch (err) {
        console.error('Erro ao excluir rascunho concluído do banco:', err);
      }
    }
    setAuditoriaAtivaId(null);
    setActiveTab('minhas_auditorias');

    notify('✓ Auditoria salva', `RAI Nº ${raiSnapshot} registrada com sucesso no banco de dados.`);
  };

  // Carrega uma auditoria do histórico no modo de VISUALIZAÇÃO (somente leitura)
  const loadAudit = async (audit, modoEditar = false) => {
    // Desvincula de qualquer rascunho que estivesse ativo ANTES de carregar os
    // dados da auditoria finalizada. Usa flushSync para garantir que o efeito
    // que reage a essa mudança (e zeraria checklist/report) termine de rodar
    // antes de continuarmos — senão ele poderia sobrescrever os dados que
    // estamos prestes a carregar (corrida de estado entre dois useEffects).
    flushSync(() => setAuditoriaAtivaId(null));

    setSelectedSector(audit.sector);
    setSelectedBranch(audit.branch);

    if (audit.checklistSnapshot && audit.reportSnapshot) {
      // Auditoria criada/editada nesta mesma sessão — já tem tudo em memória
      setChecklist(audit.checklistSnapshot);
      setReport(audit.reportSnapshot);
    } else {
      // Auditoria veio do histórico do banco (ex.: após recarregar a página) —
      // checklistSnapshot/reportSnapshot não existem localmente, então busca sob demanda
      setLoadingAuditId(audit.id);
      try {
        const itensData = await supaFetch(
          `sga_checklist_itens?auditoria_id=eq.${audit.dbId || audit.id}&select=id,ordem,pergunta,avaliacao,comentarios&order=ordem`
        );
        // Usa o ID real do banco (não uma posição sequencial) — é o mesmo ID
        // usado em sga_rncs.checklist_item_id, então preserva o vínculo com
        // RNCs já existentes ao reabrir a auditoria para edição.
        const checklistFromDb = itensData.map((it) => ({
          id: it.id,
          question: it.pergunta,
          status: it.avaliacao || '',
          comments: it.comentarios || ''
        }));
        setChecklist(checklistFromDb);
        setReport(audit._dbReport || {});
      } catch (err) {
        console.error('Erro ao carregar detalhes da auditoria:', err);
        notify('Não foi possível carregar', 'Não foi possível buscar os detalhes desta auditoria no banco. Tente novamente.');
        setLoadingAuditId(null);
        return;
      }
      setLoadingAuditId(null);
    }

    if (modoEditar && isAdmin) {
      // Modo edição: desbloqueia o checklist para a Luciene corrigir
      setChecklistStatus('Em Andamento');
      setChecklistClosedAt(null);
      setEditingAuditId(audit.dbId || audit.id);
      setAutoSaveStatus('unsaved');
      setActiveTab('checklist'); // começa pelo checklist para edição
    } else {
      setChecklistStatus('Fechado');
      setChecklistClosedAt(null);
      setEditingAuditId(null);
      setActiveTab('report');
    }
    setReopenHistory([]);
  };

  const deleteAudit = async (auditId) => {
    const audit = savedAudits.find((a) => a.id === auditId);
    setSavedAudits((prev) => prev.filter((a) => a.id !== auditId));
    // Só envia DELETE ao banco se o registro tem dbId confirmado (não é local)
    const dbId = audit?.dbId || (!audit?.isLocal ? audit?.id : null);
    if (dbId && typeof dbId === 'number') {
      try { await supaFetch(`sga_auditorias?id=eq.${dbId}`, { method: 'DELETE' }); } catch (e) {
        console.error('Erro ao excluir auditoria do banco:', e);
      }
    }
  };

  const handleChecklistChange = (id, field, value) => {
    if (isChecklistLocked) return;
    setChecklist((prev) => prev.map((item) => (item.id === id ? { ...item, [field]: value } : item)));
  };

  // handleRncChange: atualiza UI imediatamente e persiste no banco com debounce
  // (evita uma requisição por tecla digitada; aguarda 1.5s de inatividade)
  const rncSaveTimers = useRef({});
  const handleRncChange = (id, field, value) => {
    setRncs((prev) => prev.map((rnc) => (rnc.id === id ? { ...rnc, [field]: value } : rnc)));

    // Mapa de campos React → colunas do banco
    const dbFieldMap = {
      status: 'status', effective: 'eficaz', newRnc: 'necessita_nova_rnc',
      description: 'descricao_nc', correction: 'correcao',
      correctionResp: 'correcao_resp', correctionDate: 'correcao_data_limite',
      rootCause: 'causa_raiz', rootCauseResp: 'causa_raiz_resp', rootCauseDate: 'causa_raiz_data',
      actionPlan: 'plano_acao', responsible: 'plano_responsavel',
      actionPlanDatePrev: 'plano_data_prevista', actionPlanDateReal: 'plano_data_realizada',
      evidence: 'evidencias', closeDate: 'data_fechamento', closeResp: 'validador_sgq',
      process: 'processo', origin: 'origem'
    };
    const dbField = dbFieldMap[field];
    if (!dbField) return;

    // Persiste status e eficácia imediatamente (mudanças críticas)
    const immediate = field === 'status' || field === 'effective';
    const delay = immediate ? 0 : 1500;

    const rnc = rncs.find((r) => r.id === id);
    if (!rnc?.dbId) return;

    clearTimeout(rncSaveTimers.current[`${id}_${field}`]);
    rncSaveTimers.current[`${id}_${field}`] = setTimeout(async () => {
      try {
        await supaFetch(`sga_rncs?id=eq.${rnc.dbId}`, {
          method: 'PATCH',
          body: JSON.stringify({ [dbField]: value || null })
        });
      } catch (err) {
        console.error(`Erro ao salvar RNC campo ${field}:`, err);
        notify('⚠️ Erro ao salvar RNC', `Não foi possível salvar o campo "${field}" no banco. Erro: ${err.message}`);
      }
    }, delay);
  };

  // ---- Criação de RNC ----
  // Bug 3 fix: preenche todos os campos relevantes do contexto da auditoria em andamento.
  // sourceItem: item do checklist marcado NC que originou a RNC (opcional).
  const createRnc = async (sourceItem = null) => {
    const newId = pad3(rncCounter);

    // Dados da auditoria ativa
    const raiNum   = auditoriaAtiva?.raiNumber || report.raiNumber || dynamicRaiNumber;
    const auditor  = auditoriaAtiva?.report?.auditor || report.auditor || '';
    const setor    = auditoriaAtiva?.setor || selectedSector;
    const localAudId = auditoriaAtiva?.localId || null;

    // Descrição puxada automaticamente do item NC
    let descricao = '';
    if (sourceItem) {
      const itemNum = (checklist.findIndex(i => i.id === sourceItem.id) + 1) || '';
      descricao = [
        `Item ${itemNum}: ${sourceItem.question}`,
        sourceItem.comments.trim() ? `Evidência: ${sourceItem.comments.trim()}` : ''
      ].filter(Boolean).join('\n');
    }

    const newRnc = {
      id: newId,
      date: formattedDate,
      process: setor,
      origin: 'AUDITORIA INTERNA',
      sourceRaiNumber: raiNum,
      sourceAuditoriaLocalId: localAudId,
      sourceChecklistItem: sourceItem ? { id: sourceItem.id, question: sourceItem.question } : null,
      description: descricao,
      correction: '', correctionResp: auditor, correctionDate: '',
      rootCause: '', rootCauseResp: '', rootCauseDate: '',
      actionPlan: '', responsible: auditor,
      actionPlanDatePrev: '', actionPlanDateReal: '',
      evidence: '', effective: '', newRnc: '',
      closeDate: '', closeResp: '',
      status: 'Aberta'
    };
    setRncs((prev) => [newRnc, ...prev]);
    setRncCounter((prev) => prev + 1);
    setActiveTab('rnc');

    // Persiste no banco
    try {
      // Tenta encontrar o dbId da auditoria de origem: pode ser um rascunho
      // recém-salvo (localAudId) OU uma auditoria histórica sendo reaberta
      // para edição (editingAuditId) — sem esse segundo caso, RNCs criadas
      // durante uma edição ficavam sem vínculo nenhum no banco.
      const auditoriaDbId =
        savedAudits.find(a => a.id === localAudId || a.dbId === localAudId)?.dbId
        || (typeof editingAuditId === 'number' ? editingAuditId : null)
        || null;
      const payload = {
        rac_numero:       newId,
        processo:         setor,
        unidade:          selectedBranch || null,
        auditoria_id:     auditoriaDbId || null,
        origem:           'AUDITORIA INTERNA',
        descricao_nc:     descricao || null,
        correcao_resp:    auditor || null,
        plano_responsavel: auditor || null,
        status:           'Aberta'
      };
      const [created] = await supaFetch('sga_rncs', { method: 'POST', body: JSON.stringify(payload) });
      if (created?.id) {
        setRncs((prev) => prev.map((r) => r.id === newId ? { ...r, dbId: created.id } : r));
      }
    } catch (err) {
      console.error('Erro ao criar RNC no banco:', err);
      notify('⚠️ Erro ao criar RNC', `RNC criada localmente mas não registrada no banco. Erro: ${err.message}`);
    }
  };

  // ---- Calendário de programação de auditorias ----

  const buildAuditEmail = (event) => {
    const dateObj = dateKeyToDate(event.dateKey);
    const dateLabel = dateObj.toLocaleDateString('pt-BR', { day: '2-digit', month: '2-digit', year: 'numeric', weekday: 'long' });
    const timeRange = formatTimeRange(event.time, event.durationMin);
    return {
      subject: `Programação de Auditoria Interna — ${event.sector} — ${dateObj.toLocaleDateString('pt-BR')}`,
      sector: event.sector,
      body: `Prezados,

Informamos a programação de auditoria interna do Sistema de Gestão da Qualidade para o setor ${event.sector}.

Data: ${dateLabel}
Horário: ${timeRange}
Unidade: ${event.branch}
Auditor Líder: ${event.auditor}

Solicitamos a disponibilidade do responsável pelo setor e da documentação pertinente ao processo no horário indicado.

Qualquer necessidade de reagendamento, favor retornar com antecedência.

Atenciosamente,
Gestão da Qualidade — Kalenborn do Brasil`
    };
  };

  const openEventEmail = (event) => {
    setEmailCopied(false);
    setEmailDraft(buildAuditEmail(event));
  };

  // Persiste um evento no banco em background, sem bloquear a UI
  const persistEvent = async (event, isNew) => {
    try {
      const setorRow = sectors.find((s) => s.nome === event.sector);
      const auditorRow = auditoresDb.find((a) => a.nome === event.auditor);
      const payload = {
        setor_id: setorRow?.id || null,
        auditor_id: auditorRow?.id || null,
        titulo: event.title || `Auditoria ${event.sector || ''}`.trim(),
        responsavel_nome: event.auditor || null,
        data_auditoria: event.dateKey,
        horario: event.time + ':00',
        duracao_min: event.durationMin,
        unidade: event.branch,
        observacoes: event.notes || null,
        status: 'Programado'
      };
      if (isNew) {
        const [created] = await supaFetch('sga_eventos_calendario', { method: 'POST', body: JSON.stringify(payload) });
        return created?.id;
      } else if (event.dbId) {
        await supaFetch(`sga_eventos_calendario?id=eq.${event.dbId}`, { method: 'PATCH', body: JSON.stringify(payload) });
      }
    } catch (err) {
      console.error('Erro ao persistir evento:', err);
      notify('⚠️ Erro ao salvar evento', `Evento salvo localmente mas não registrado no banco. Erro: ${err.message}`);
    }
    return null;
  };

  const saveEvent = async (eventData) => {
    const isNew = !eventData.id || typeof eventData.id === 'number' && eventData.id > 4999; // IDs locais começam em 5001
    const localId = isNew ? nextEventId() : eventData.id;
    const finalEvent = { ...eventData, id: localId };
    setAuditEvents((prev) => isNew ? [...prev, finalEvent] : prev.map((e) => (e.id === finalEvent.id ? finalEvent : e)));
    setEventDraft(null);
    openEventEmail(finalEvent);
    const dbId = await persistEvent(finalEvent, isNew);
    if (dbId) setAuditEvents((prev) => prev.map((e) => e.id === localId ? { ...e, dbId } : e));
  };

  const deleteEvent = async (eventId) => {
    const ev = auditEvents.find((e) => e.id === eventId);
    setAuditEvents((prev) => prev.filter((e) => e.id !== eventId));
    if (ev?.dbId) {
      try { await supaFetch(`sga_eventos_calendario?id=eq.${ev.dbId}`, { method: 'DELETE' }); } catch {}
    }
  };

  const handleDropOnDate = async (dateKey) => {
    if (!draggedEventId) return;
    const moved = auditEvents.find((e) => e.id === draggedEventId);
    setDraggedEventId(null);
    setDragOverDateKey(null);
    if (!moved || moved.dateKey === dateKey) return;
    const updated = { ...moved, dateKey };
    setAuditEvents((prev) => prev.map((e) => (e.id === updated.id ? updated : e)));
    openEventEmail(updated);
    if (updated.dbId) {
      try { await supaFetch(`sga_eventos_calendario?id=eq.${updated.dbId}`, { method: 'PATCH', body: JSON.stringify({ data_auditoria: dateKey }) }); } catch {}
    }
  };

  const getNcCount = () => checklist.filter((item) => item.status === 'NC').length;
  const getCCount = () => checklist.filter((item) => item.status === 'C').length;
  const getObsCount = () => checklist.filter((item) => item.status === 'Obs').length;
  const getAnsweredCount = () => checklist.filter((item) => item.status !== '').length;
  const getProgressPct = () => (checklist.length === 0 ? 0 : Math.round((getAnsweredCount() / checklist.length) * 100));

  const getCompiledObservations = () => {
    const obsItems = checklist
      .map((item, idx) => ({ ...item, num: idx + 1 }))
      .filter((item) => item.status === 'Obs' && item.comments.trim() !== '');
    if (obsItems.length === 0) return report.observations || '';
    return obsItems
      .map((item, i) => `${i + 1} – Ref. item ${item.num}: ${item.comments.trim()}`)
      .join('\n\n');
  };

  useEffect(() => {
    // Só recalcula automaticamente enquanto o checklist está editável (rascunho
    // ativo ou em modo de edição). Em modo de VISUALIZAÇÃO (checklistStatus ===
    // 'Fechado' e não está em edição), isso sobrescrevia o RAI e as observações
    // salvos pelos valores do rascunho atualmente em aberto — mostrando dados
    // errados na tela de "VER".
    if (activeTab === 'report' && checklistStatus !== 'Fechado') {
      const compiledObs = getCompiledObservations();
      const newObsText = checklist.some((i) => i.status === 'Obs' && i.comments) ? compiledObs : report.observations;
      setReport((prev) => ({ ...prev, observations: newObsText, raiNumber: dynamicRaiNumber }));
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activeTab, checklist, dynamicRaiNumber, checklistStatus]);

  // Itens NC do checklist ativo que ainda não têm RNC vinculada
  const ncItemsWithoutRnc = useMemo(() => {
    const raiNum = auditoriaAtiva?.raiNumber || report.raiNumber;
    const linkedIds = new Set(
      rncs.filter((r) => r.sourceChecklistItem)
          .map((r) => `${r.sourceRaiNumber}::${r.sourceChecklistItem.id}`)
    );
    return checklist.filter(
      (item) => item.status === 'NC' && !linkedIds.has(`${raiNum}::${item.id}`)
    );
  }, [checklist, rncs, auditoriaAtiva, report.raiNumber]);

  // Eventos do calendário a partir de hoje, ordenados por data/hora — usados
  // no card "Próximas Auditorias Programadas" do dashboard.
  const upcomingEvents = useMemo(() => {
    const todayKey = toDateKey(new Date());
    return auditEvents
      .filter((ev) => ev.dateKey >= todayKey)
      .sort((a, b) => (a.dateKey === b.dateKey ? a.time.localeCompare(b.time) : a.dateKey.localeCompare(b.dateKey)));
  }, [auditEvents]);

  // window.print() é bloqueado silenciosamente dentro de iframes sandboxed
  // (ex: previews de ferramentas de IA como Gemini/Claude/v0). Para esses casos,
  // tentamos abrir uma nova janela com o conteúdo e disparar a impressão lá,
  // que tem permissão própria. Se o navegador bloquear o pop-up, avisamos o usuário.
  const printDocument = () => {
    try {
      const inIframe = window.self !== window.top;
      if (!inIframe) {
        window.print();
        return;
      }

      const printable = document.querySelector('.print-area') || document.querySelector('main');
      if (!printable) {
        window.print();
        return;
      }

      // innerHTML não captura o valor atual digitado em <input>/<textarea>/<select>
      // (só o valor inicial do DOM), então sincronizamos manualmente antes de clonar.
      const clone = printable.cloneNode(true);
      const liveFields = printable.querySelectorAll('input, textarea, select');
      const clonedFields = clone.querySelectorAll('input, textarea, select');
      liveFields.forEach((field, i) => {
        const target = clonedFields[i];
        if (!target) return;
        if (field.tagName === 'TEXTAREA') {
          target.textContent = field.value;
          target.value = field.value;
        } else if (field.tagName === 'SELECT') {
          target.value = field.value;
          Array.from(target.options).forEach((opt) => {
            opt.selected = opt.value === field.value;
          });
        } else if (field.type === 'checkbox' || field.type === 'radio') {
          if (field.checked) target.setAttribute('checked', 'checked');
          else target.removeAttribute('checked');
        } else {
          target.setAttribute('value', field.value);
          target.value = field.value;
        }
      });

      const printWindow = window.open('', '_blank', 'width=1000,height=800');
      if (!printWindow) {
        notify('Pop-up bloqueado', 'O navegador bloqueou a janela de impressão. Permita pop-ups para este site, ou abra o sistema fora do preview (em uma aba normal do navegador) e tente novamente.');
        return;
      }

      const styles = Array.from(document.querySelectorAll('link[rel="stylesheet"], style'))
        .map((node) => node.outerHTML)
        .join('\n');

      printWindow.document.write(`
        <!DOCTYPE html>
        <html>
          <head>
            <meta charset="utf-8" />
            <title>Impressão — Sistema de Auditoria</title>
            ${styles}
            <style>
              body { margin: 0; padding: 24px; background: #fff; }
              input, textarea, select { font-family: inherit; }
              @media print { body { padding: 0; } }
            </style>
          </head>
          <body>${clone.innerHTML}</body>
        </html>
      `);
      printWindow.document.close();

      printWindow.onload = () => {
        printWindow.focus();
        printWindow.print();
      };
    } catch (err) {
      notify('Não foi possível imprimir', 'A impressão não está disponível neste ambiente de preview. Abra o sistema publicado em uma aba normal do navegador (fora do preview da ferramenta de IA) para imprimir ou gerar PDF.');
    }
  };

  // ---- Exportação para Excel (.xlsx) via SheetJS carregado por CDN ----
  const ensureXlsxOrWarn = () => {
    if (typeof window !== 'undefined' && XLSX) return true;
    notify('Preparando exportação', 'A ferramenta de exportação para Excel ainda está carregando. Aguarde alguns segundos e tente novamente.');
    return false;
  };

  const downloadWorkbook = (wb, filename) => {
    XLSX.writeFile(wb, filename);
  };

  // ---- Helper para aplicar estilos de célula (borda + negrito + alinhamento) ----
  const xlsxStyle = (bold = false, bg = null, align = 'left', wrapText = true) => ({
    font: { bold, name: 'Arial', sz: 10 },
    alignment: { horizontal: align, vertical: 'center', wrapText },
    border: {
      top:    { style: 'thin', color: { rgb: '000000' } },
      bottom: { style: 'thin', color: { rgb: '000000' } },
      left:   { style: 'thin', color: { rgb: '000000' } },
      right:  { style: 'thin', color: { rgb: '000000' } }
    },
    fill: bg ? { fgColor: { rgb: bg }, patternType: 'solid' } : undefined
  });

  const xlsxHeader = (bold = true) => xlsxStyle(bold, 'D9D9D9', 'center');

  // Aplica estilo a um range de células
  const styleRange = (ws, startRow, startCol, endRow, endCol, style) => {
    for (let r = startRow; r <= endRow; r++) {
      for (let c = startCol; c <= endCol; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        if (ws[addr]) ws[addr].s = style;
      }
    }
  };

  // Celula com valor + estilo
  const cell = (v, s) => ({ v, t: typeof v === 'number' ? 'n' : 's', s });

  // ---- CHECKLIST — layout visual fiel ao formulário KdB147 ----
  const exportChecklistExcel = () => {
    if (!ensureXlsxOrWarn()) return;
    const X = XLSX;
    const aoa = [
      // Linha 1: Logo + Título + Setor/Data
      [cell('Kalenborn', xlsxHeader()), cell(''), cell('KdB147 - LISTA DE VERIFICAÇÃO DE AUDITORIA INTERNA', xlsxHeader()), cell(''), cell('SETOR', xlsxHeader()), cell(selectedSector, xlsxStyle(true, null, 'center'))],
      [cell('Wear Protection Solutions', xlsxStyle(false, 'D9D9D9', 'left')), cell(''), cell(''), cell(''), cell('DATA', xlsxHeader()), cell(formattedDate, xlsxStyle(false, null, 'center'))],
      // Linha 3: Requisitos
      [cell('Requisitos ISO:', xlsxHeader()), cell(''), cell(selectedSectorRequisitos, xlsxStyle(false, null, 'left')), cell(''), cell(''), cell('')],
      // Linha 4: Documentos
      [cell('Documentos:', xlsxHeader()), cell(''), cell('Norma NBR ISO 9001:2015 / Manual SGQ / Procedimentos / Mapa de Processo / Formulários', xlsxStyle(false, null, 'left')), cell(''), cell(''), cell('')],
      // Linha 5: Legenda
      [cell('Legenda:', xlsxHeader()), cell('C - Conforme', xlsxStyle(false, 'CCFFCC', 'center')), cell('NC - Não Conforme', xlsxStyle(false, 'FFCCCC', 'center')), cell('Obs - Observação', xlsxStyle(false, 'CCE5FF', 'center')), cell(''), cell('')],
      // Linha 6: Cabeçalho da tabela
      [cell('Nº', xlsxHeader()), cell('Questionamento Operacional', xlsxHeader()), cell(''), cell('Evidências / Comentários', xlsxHeader()), cell(''), cell('Avaliação', xlsxHeader())],
      // Itens do checklist
      ...checklist.map((item, idx) => [
        cell(idx + 1, xlsxStyle(true, 'F2F2F2', 'center')),
        cell(item.question, xlsxStyle(false, null, 'left')),
        cell('', xlsxStyle()),
        cell(item.comments, xlsxStyle(false, 'FFFDE7', 'left')),
        cell('', xlsxStyle()),
        cell(item.status || '', xlsxStyle(true,
          item.status === 'C' ? 'CCFFCC' : item.status === 'NC' ? 'FFCCCC' : item.status === 'Obs' ? 'CCE5FF' : null,
          'center'))
      ])
    ];

    const ws = X.utils.aoa_to_sheet(aoa);

    // Mesclagens
    ws['!merges'] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }, // Logo linha 1
      { s: { r: 1, c: 0 }, e: { r: 1, c: 1 } }, // Logo linha 2
      { s: { r: 0, c: 2 }, e: { r: 1, c: 3 } }, // Título
      { s: { r: 2, c: 0 }, e: { r: 2, c: 1 } }, // Req label
      { s: { r: 2, c: 2 }, e: { r: 2, c: 5 } }, // Req value
      { s: { r: 3, c: 0 }, e: { r: 3, c: 1 } }, // Docs label
      { s: { r: 3, c: 2 }, e: { r: 3, c: 5 } }, // Docs value
      { s: { r: 4, c: 4 }, e: { r: 4, c: 5 } }, // Legenda extra
      { s: { r: 5, c: 1 }, e: { r: 5, c: 2 } }, // Header pergunta
      { s: { r: 5, c: 3 }, e: { r: 5, c: 4 } }, // Header evidência
      ...checklist.map((_, idx) => ({ s: { r: 6 + idx, c: 1 }, e: { r: 6 + idx, c: 2 } })),
      ...checklist.map((_, idx) => ({ s: { r: 6 + idx, c: 3 }, e: { r: 6 + idx, c: 4 } })),
    ];

    ws['!cols'] = [{ wch: 6 }, { wch: 40 }, { wch: 10 }, { wch: 40 }, { wch: 10 }, { wch: 14 }];
    ws['!rows'] = [{ hpt: 24 }, { hpt: 18 }, { hpt: 24 }, { hpt: 24 }, { hpt: 24 }, { hpt: 20 },
      ...checklist.map(() => ({ hpt: 40 }))];

    const wb = X.utils.book_new();
    X.utils.book_append_sheet(wb, ws, 'Checklist');
    downloadWorkbook(wb, `Checklist_${selectedSector.replace(/[^a-zA-Z0-9]/g, '_')}_${formattedDate.replace(/\//g, '-')}.xlsx`);
  };

  // ---- RELATÓRIO RAI — layout visual fiel ao formulário ----
  const exportReportExcel = () => {
    if (!ensureXlsxOrWarn()) return;
    const X = XLSX;
    const lbl = (v) => cell(v, xlsxStyle(true, 'D9D9D9', 'left'));
    const val = (v) => cell(v || '', xlsxStyle(false, null, 'left'));
    const title = (v) => cell(v, xlsxStyle(true, 'BFBFBF', 'center'));

    const aoa = [
      [cell('Kalenborn — Wear Protection Solutions', xlsxStyle(true, 'D9D9D9', 'center')), cell(''), cell('RELATÓRIO DE AUDITORIA INTERNA', xlsxStyle(true, 'D9D9D9', 'center')), cell(''), cell('RAI Nº', xlsxHeader()), cell(report.raiNumber, xlsxStyle(true, 'FFE0E0', 'center'))],
      [lbl('Auditoria Inicial'), val(formattedDate), lbl('Data de Emissão'), val(report.date), lbl('Unidade'), val(selectedBranch)],
      [lbl('Área'), cell(selectedSector, xlsxStyle(true, null, 'left')), cell(''), cell(''), cell(''), cell('')],
      [lbl('Processo'), val(selectedSector), cell(''), cell(''), cell(''), cell('')],
      [lbl('Metodologia'), cell('Entrevista, Amostragem e Análise de Documentos do SGQ', xlsxStyle(false, null, 'left')), cell(''), cell(''), cell(''), cell('')],
      [lbl('Objetivo'), cell('Verificar a conformidade e eficácia do Processo dentro do SGQ Kalenborn, alinhado à ISO 9001:2015.', xlsxStyle(false, null, 'left')), cell(''), cell(''), cell(''), cell('')],
      [lbl('Auditor Líder'), val(report.auditor), cell(''), cell(''), cell(''), cell('')],
      [lbl('Auditado'), val(report.auditee), cell(''), cell(''), cell(''), cell('')],
      [title('PONTOS FORTES / POSITIVOS IDENTIFICADOS'), cell(''), cell(''), cell(''), cell(''), cell('')],
      [cell(report.positivePoints || '', xlsxStyle(false, 'F9FFF9', 'left', true)), cell(''), cell(''), cell(''), cell(''), cell('')],
      [title('OBSERVAÇÕES PERTINENTES'), cell(''), cell(''), cell(''), cell(''), cell('')],
      [cell(report.observations || '', xlsxStyle(false, 'FFFDE7', 'left', true)), cell(''), cell(''), cell(''), cell(''), cell('')],
      [title('OPORTUNIDADES E SUGESTÕES DE MELHORIA'), cell(''), cell(''), cell(''), cell(''), cell('')],
      [cell(report.improvements || '', xlsxStyle(false, 'F0F8FF', 'left', true)), cell(''), cell(''), cell(''), cell(''), cell('')],
      [lbl('QUANTIDADE DE RNCs EXIGIDAS'), cell(''), cell(''), cell(''), cell(''), cell(getNcCount(), xlsxStyle(true, 'FFE0E0', 'center'))],
      [title('CONCLUSÃO OFICIAL DA AUDITORIA'), cell(''), cell(''), cell(''), cell(''), cell('')],
      [cell(report.conclusion || '', xlsxStyle(false, null, 'center', true)), cell(''), cell(''), cell(''), cell(''), cell('')],
      [lbl('Assinatura do Auditor Líder'), cell(''), cell(''), lbl('Responsável SGQ Unidade'), cell(''), cell('')],
      [cell(report.auditor || '', xlsxStyle(false, null, 'center')), cell(''), cell(''), val(selectedBranch), cell(''), cell('')],
    ];

    const ws = X.utils.aoa_to_sheet(aoa);
    ws['!merges'] = [
      { s:{r:0,c:0}, e:{r:0,c:1} }, { s:{r:0,c:2}, e:{r:0,c:3} },
      ...[[2],[3],[4],[5],[6],[7]].map(([r]) => ({ s:{r,c:1}, e:{r,c:5} })),
      ...[8,10,12,14,15,16].map(r => ({ s:{r,c:0}, e:{r,c:5} })),
      { s:{r:9,c:0}, e:{r:9,c:5} }, { s:{r:11,c:0}, e:{r:11,c:5} },
      { s:{r:13,c:0}, e:{r:13,c:5} }, { s:{r:16,c:0}, e:{r:16,c:5} },
      { s:{r:17,c:0}, e:{r:17,c:2} }, { s:{r:17,c:3}, e:{r:17,c:5} },
      { s:{r:18,c:0}, e:{r:18,c:2} }, { s:{r:18,c:3}, e:{r:18,c:5} },
    ];
    ws['!cols'] = [{wch:20},{wch:20},{wch:20},{wch:20},{wch:10},{wch:16}];
    ws['!rows'] = [{hpt:24},{hpt:18},{hpt:20},{hpt:20},{hpt:20},{hpt:36},{hpt:20},{hpt:20},
      {hpt:18},{hpt:60},{hpt:18},{hpt:60},{hpt:18},{hpt:60},{hpt:24},{hpt:18},{hpt:60},{hpt:24},{hpt:40}];

    const wb = X.utils.book_new();
    X.utils.book_append_sheet(wb, ws, 'Relatorio RAI');
    downloadWorkbook(wb, `RAI_${report.raiNumber.replace('/', '-')}.xlsx`);
  };

  // ---- RNC — cada RAC em aba separada ----
  const exportRncExcel = () => {
    if (!ensureXlsxOrWarn()) return;
    const X = XLSX;
    const wb = X.utils.book_new();
    const lbl = (v) => cell(v, xlsxStyle(true, 'D9D9D9', 'left'));
    const val = (v) => cell(v || '', xlsxStyle(false, null, 'left'));
    const sec = (v) => cell(v, xlsxStyle(true, 'BFBFBF', 'center'));

    const rncList = rncs.length ? rncs : [{ id: '—', date: '—', process: '', origin: '', description: '', correction: '', correctionResp: '', correctionDate: '', rootCause: '', rootCauseResp: '', rootCauseDate: '', actionPlan: '', responsible: '', actionPlanDatePrev: '', actionPlanDateReal: '', evidence: '', effective: '', newRnc: '', closeDate: '', closeResp: '', status: '' }];

    rncList.forEach((rnc) => {
      const aoa = [
        [cell('Kalenborn — Wear Protection Solutions', xlsxStyle(true,'D9D9D9','center')), cell(''), cell('RELATÓRIO DE AÇÕES CORRETIVAS (RAC)', xlsxStyle(true,'D9D9D9','center')), cell(''), cell('Nº R.A.C', xlsxHeader()), cell(rnc.id, xlsxStyle(true,'FFE0E0','center'))],
        [lbl('Processo'), cell(rnc.process||'', xlsxStyle(true,null,'left')), cell(''), lbl('Data de Emissão'), val(rnc.date), cell('')],
        [lbl('Origem'), cell(rnc.origin||'AUDITORIA INTERNA', xlsxStyle(false,null,'left')), cell(''), lbl('Unidade'), val(selectedBranch), cell('')],
        [sec('DESCRIÇÃO DA NÃO CONFORMIDADE'), cell(''), cell(''), cell(''), cell(''), cell('')],
        [cell(rnc.description||'', xlsxStyle(false,'FFF0F0','left',true)), cell(''), cell(''), cell(''), cell(''), cell('')],
        [sec('AÇÃO DE CORREÇÃO IMEDIATA'), cell(''), lbl('Responsável'), val(rnc.correctionResp), lbl('Data Limite'), val(rnc.correctionDate)],
        [cell(rnc.correction||'', xlsxStyle(false,null,'left',true)), cell(''), cell(''), cell(''), cell(''), cell('')],
        [sec('ANÁLISE DE CAUSA RAIZ'), cell(''), lbl('Líder da Análise'), val(rnc.rootCauseResp), lbl('Data Conclusão'), val(rnc.rootCauseDate)],
        [cell(rnc.rootCause||'', xlsxStyle(false,'FFF5F5','left',true)), cell(''), cell(''), cell(''), cell(''), cell('')],
        [sec('PLANO DE AÇÃO'), cell(''), lbl('Responsável'), val(rnc.responsible), lbl('Previsto'), val(rnc.actionPlanDatePrev)],
        [cell(rnc.actionPlan||'', xlsxStyle(false,null,'left',true)), cell(''), cell(''), cell(''), lbl('Realizado'), val(rnc.actionPlanDateReal)],
        [sec('EVIDÊNCIAS E ANÁLISE DE EFICÁCIA'), cell(''), cell(''), cell(''), cell(''), cell('')],
        [cell(rnc.evidence||'', xlsxStyle(false,'F9FFF9','left',true)), cell(''), cell(''), cell(''), cell(''), cell('')],
        [lbl('Ação foi Eficaz?'), cell(rnc.effective||'', xlsxStyle(true,null,'center')), cell(''), lbl('Necessita Nova RNC?'), cell(rnc.newRnc||'', xlsxStyle(true,null,'center')), cell('')],
        [lbl('Data Fechamento'), val(rnc.closeDate), cell(''), lbl('Validador SGQ'), val(rnc.closeResp), cell('')],
        [lbl('Status'), cell(rnc.status||'Aberta', xlsxStyle(true, rnc.status==='Fechada'?'CCFFCC':'FFE0E0','center')), cell(''), cell(''), cell(''), cell('')],
      ];

      const ms = [
        {s:{r:0,c:0},e:{r:0,c:1}}, {s:{r:0,c:2},e:{r:0,c:3}},
        {s:{r:1,c:1},e:{r:1,c:2}}, {s:{r:2,c:1},e:{r:2,c:2}},
        ...[3,7,9,11].map(r=>({s:{r,c:0},e:{r,c:5}})),
        ...[4,8,12].map(r=>({s:{r,c:0},e:{r,c:5}})),
        {s:{r:5,c:0},e:{r:5,c:1}}, {s:{r:6,c:0},e:{r:6,c:1}},
        {s:{r:7,c:0},e:{r:7,c:1}}, {s:{r:9,c:0},e:{r:9,c:1}},
        {s:{r:10,c:0},e:{r:10,c:3}}, {s:{r:10,c:4},e:{r:10,c:5}},
        {s:{r:13,c:0},e:{r:13,c:2}}, {s:{r:13,c:3},e:{r:13,c:5}},
        {s:{r:14,c:0},e:{r:14,c:2}}, {s:{r:14,c:3},e:{r:14,c:5}},
        {s:{r:15,c:1},e:{r:15,c:5}},
      ];

      const ws = X.utils.aoa_to_sheet(aoa);
      ws['!merges'] = ms;
      ws['!cols'] = [{wch:18},{wch:18},{wch:10},{wch:18},{wch:14},{wch:14}];
      ws['!rows'] = [{hpt:24},{hpt:20},{hpt:20},{hpt:18},{hpt:60},{hpt:18},{hpt:60},{hpt:18},{hpt:60},{hpt:18},{hpt:60},{hpt:18},{hpt:60},{hpt:20},{hpt:20},{hpt:20}];
      const sheetName = `RAC_${String(rnc.id).substring(0,10)}`;
      X.utils.book_append_sheet(wb, ws, sheetName);
    });

    downloadWorkbook(wb, `RNC_${selectedBranch.replace(/[^a-zA-Z0-9]/g,'_')}_${formattedDate.replace(/\//g,'-')}.xlsx`);
  };

  // ---- HISTÓRICO — tabela formatada ----
  const exportHistoryExcel = () => {
    if (!ensureXlsxOrWarn()) return;
    const X = XLSX;
    const hdr = (v) => cell(v, xlsxStyle(true,'D9D9D9','center'));
    const aoa = [
      [cell('Kalenborn do Brasil — Repositório de Auditorias Internas ISO 9001:2015', xlsxStyle(true,'BFBFBF','center')), cell(''), cell(''), cell(''), cell(''), cell(''), cell('')],
      [hdr('Nº RAI'), hdr('Data'), hdr('Unidade'), hdr('Setor / Processo'), hdr('Auditor Líder'), hdr('Qtd. NCs'), hdr('Status')],
      ...savedAudits.map(a => [
        cell(a.raiNumber, xlsxStyle(true,null,'center')),
        cell(a.date, xlsxStyle(false,null,'center')),
        cell(a.branch, xlsxStyle(false,null,'left')),
        cell(a.sector, xlsxStyle(false,null,'left')),
        cell(a.auditor, xlsxStyle(false,null,'left')),
        cell(a.ncCount, xlsxStyle(true, a.ncCount>0?'FFCCCC':'CCFFCC','center')),
        cell(a.status, xlsxStyle(false,null,'center')),
      ])
    ];

    const ws = X.utils.aoa_to_sheet(aoa);
    ws['!merges'] = [{ s:{r:0,c:0}, e:{r:0,c:6} }];
    ws['!cols'] = [{wch:12},{wch:12},{wch:26},{wch:32},{wch:22},{wch:10},{wch:14}];
    ws['!rows'] = [{hpt:28},{hpt:20},...savedAudits.map(()=>({hpt:18}))];

    const wb = X.utils.book_new();
    X.utils.book_append_sheet(wb, ws, 'Historico');
    downloadWorkbook(wb, `Historico_Auditorias_${formattedDate.replace(/\//g,'-')}.xlsx`);
  };



  // ---- Métricas globais derivadas do estado real ----
  const totalAuditsYear = savedAudits.length;
  const totalOpenRncs = rncs.filter((r) => r.status === 'Aberta').length;
  const auditsWithZeroNc = savedAudits.filter((a) => a.ncCount === 0).length;
  const conformityIndex = savedAudits.length > 0
    ? Math.round((auditsWithZeroNc / savedAudits.length) * 100)
    : 100;
  const effectivenessIndex = (() => {
    const closed = rncs.filter((r) => r.effective === 'Sim').length;
    const evaluated = rncs.filter((r) => r.effective === 'Sim' || r.effective === 'Não').length;
    return evaluated > 0 ? Math.round((closed / evaluated) * 100) : null;
  })();

  // ---- Filtro e paginação do histórico ----
  const filteredAudits = useMemo(() => {
    const term = historySearch.trim().toLowerCase();
    return savedAudits.filter((a) => {
      const matchesTerm = !term ||
        a.raiNumber.toLowerCase().includes(term) ||
        a.sector.toLowerCase().includes(term) ||
        a.auditor.toLowerCase().includes(term) ||
        a.branch.toLowerCase().includes(term);
      const matchesNc = historyNcFilter === 'Todas' ||
        (historyNcFilter === 'ComNC' && a.ncCount > 0) ||
        (historyNcFilter === 'SemNC' && a.ncCount === 0);
      return matchesTerm && matchesNc;
    });
  }, [savedAudits, historySearch, historyNcFilter]);

  const historyTotalPages = Math.max(1, Math.ceil(filteredAudits.length / HISTORY_PAGE_SIZE));
  const pagedAudits = filteredAudits.slice(
    (historyPage - 1) * HISTORY_PAGE_SIZE,
    historyPage * HISTORY_PAGE_SIZE
  );

  useEffect(() => {
    setHistoryPage(1);
  }, [historySearch]);

  // ---- RENDERIZADORES DE ABAS ----

  const renderDashboard = () => (
    <div className="space-y-6 animate-fade-in">
      <div className="flex flex-col sm:flex-row sm:justify-between sm:items-center gap-4 mb-2">
        <div>
          <p className="text-[11px] font-bold text-indigo-600 uppercase tracking-[0.2em] mb-1">Sistema de Gestão de Auditoria · ISO 9001:2015</p>
          <h2 className="text-3xl font-black text-slate-900 tracking-tight">Painel Executivo</h2>
          <p className="text-slate-500 mt-1">Visão consolidada — {selectedBranch}</p>
        </div>
        <button
          onClick={() => setActiveTab('checklist')}
          className="bg-indigo-600 hover:bg-indigo-700 text-white px-5 py-3 rounded-xl font-bold transition shadow-lg shadow-indigo-600/25 flex items-center gap-2 self-start"
        >
          <Plus size={18} strokeWidth={2.5} /> Iniciar Nova Auditoria
        </button>
      </div>

      {/* KPI Cards — clicáveis: levam direto para a tela e filtro relacionados */}
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-5">
        <button
          onClick={() => setActiveTab('gestao')}
          className="text-left bg-white p-6 rounded-2xl shadow-sm border border-slate-100 relative overflow-hidden group hover:shadow-md hover:border-indigo-200 transition-all cursor-pointer"
        >
          <div className="absolute -right-6 -top-6 w-24 h-24 bg-indigo-50 rounded-full group-hover:scale-110 transition-transform"></div>
          <div className="relative flex justify-between items-start mb-5">
            <h3 className="text-xs font-bold text-slate-400 uppercase tracking-wider">Auditorias no Ano</h3>
            <div className="p-2.5 bg-indigo-600 text-white rounded-xl shadow-md shadow-indigo-600/30"><ClipboardCheck size={18} /></div>
          </div>
          <div className="relative text-4xl font-black text-slate-900 font-mono tabular-nums">{totalAuditsYear}</div>
          <div className="relative text-xs text-slate-500 font-medium mt-2 flex items-center gap-1">
            Registradas em {currentYear} <ChevronRight size={12} className="opacity-0 group-hover:opacity-100 transition" />
          </div>
        </button>

        <button
          onClick={() => { setRncStatusFilter('Aberta'); setActiveTab('rnc'); }}
          className="text-left bg-white p-6 rounded-2xl shadow-sm border border-slate-100 relative overflow-hidden group hover:shadow-md hover:border-rose-200 transition-all cursor-pointer"
        >
          <div className="absolute -right-6 -top-6 w-24 h-24 bg-rose-50 rounded-full group-hover:scale-110 transition-transform"></div>
          <div className="relative flex justify-between items-start mb-5">
            <h3 className="text-xs font-bold text-slate-400 uppercase tracking-wider">RNCs Abertas</h3>
            <div className="p-2.5 bg-rose-600 text-white rounded-xl shadow-md shadow-rose-600/30"><AlertTriangle size={18} /></div>
          </div>
          <div className="relative text-4xl font-black text-slate-900 font-mono tabular-nums">{totalOpenRncs}</div>
          <div className="relative text-xs font-medium mt-2 text-rose-600 flex items-center gap-1">
            {totalOpenRncs > 0 ? 'Requer atenção imediata' : 'Nenhuma pendência'} <ChevronRight size={12} className="opacity-0 group-hover:opacity-100 transition" />
          </div>
        </button>

        <button
          onClick={() => { setHistoryNcFilter('SemNC'); setActiveTab('gestao'); }}
          className="text-left bg-white p-6 rounded-2xl shadow-sm border border-slate-100 relative overflow-hidden group hover:shadow-md hover:border-emerald-200 transition-all cursor-pointer"
        >
          <div className="absolute -right-6 -top-6 w-24 h-24 bg-emerald-50 rounded-full group-hover:scale-110 transition-transform"></div>
          <div className="relative flex justify-between items-start mb-5">
            <h3 className="text-xs font-bold text-slate-400 uppercase tracking-wider">Índice de Conformidade</h3>
            <div className="p-2.5 bg-emerald-600 text-white rounded-xl shadow-md shadow-emerald-600/30"><ShieldCheck size={18} /></div>
          </div>
          <div className="relative text-4xl font-black text-slate-900 font-mono tabular-nums">{conformityIndex}%</div>
          <div className="relative text-xs text-slate-500 font-medium mt-2 flex items-center gap-1">
            Meta: 90.0% · Auditorias sem NC <ChevronRight size={12} className="opacity-0 group-hover:opacity-100 transition" />
          </div>
        </button>

        <button
          onClick={() => { setRncStatusFilter('Todas'); setActiveTab('rnc'); }}
          className="text-left bg-white p-6 rounded-2xl shadow-sm border border-slate-100 relative overflow-hidden group hover:shadow-md hover:border-violet-200 transition-all cursor-pointer"
        >
          <div className="absolute -right-6 -top-6 w-24 h-24 bg-violet-50 rounded-full group-hover:scale-110 transition-transform"></div>
          <div className="relative flex justify-between items-start mb-5">
            <h3 className="text-xs font-bold text-slate-400 uppercase tracking-wider">Eficácia de Ações</h3>
            <div className="p-2.5 bg-violet-600 text-white rounded-xl shadow-md shadow-violet-600/30"><Gauge size={18} /></div>
          </div>
          <div className="relative text-4xl font-black text-slate-900 font-mono tabular-nums">
            {effectivenessIndex !== null ? `${effectivenessIndex}%` : '—'}
          </div>
          <div className="relative text-xs text-slate-500 font-medium mt-2 flex items-center gap-1">
            {effectivenessIndex !== null ? 'Ações validadas e fechadas' : 'Ainda sem ações avaliadas'} <ChevronRight size={12} className="opacity-0 group-hover:opacity-100 transition" />
          </div>
        </button>
      </div>


      {/* Histórico geral — agrega TODAS as auditorias já concluídas (não só a atual) */}
      <div className="bg-white rounded-2xl shadow-sm border border-slate-100 p-6">
        <div className="flex justify-between items-center mb-5">
          <div>
            <h3 className="text-base font-bold text-slate-800">Histórico Geral de Conformidade</h3>
            <p className="text-slate-500 text-xs mt-0.5">Somatório de todos os itens de checklist de todas as auditorias já concluídas</p>
          </div>
          <button onClick={() => setActiveTab('gestao')} className="text-xs text-indigo-600 font-bold hover:underline shrink-0">Ver Histórico Completo</button>
        </div>
        {historicalStats.total === 0 ? (
          <div className="p-6 text-center text-slate-400 text-sm font-medium">Nenhum item de checklist concluído ainda.</div>
        ) : (
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
            {[
              { label: 'Conforme',     count: historicalStats.c,   classes: {
                  card: 'rounded-xl border border-emerald-100 bg-emerald-50/50 p-5',
                  count: 'text-3xl font-black font-mono text-emerald-600',
                  pct: 'text-lg font-black font-mono text-emerald-500',
                  bar: 'h-full bg-emerald-500 rounded-full transition-all'
              } },
              { label: 'Não Conforme', count: historicalStats.nc,  classes: {
                  card: 'rounded-xl border border-rose-100 bg-rose-50/50 p-5',
                  count: 'text-3xl font-black font-mono text-rose-600',
                  pct: 'text-lg font-black font-mono text-rose-500',
                  bar: 'h-full bg-rose-500 rounded-full transition-all'
              } },
              { label: 'Observação',   count: historicalStats.obs, classes: {
                  card: 'rounded-xl border border-amber-100 bg-amber-50/50 p-5',
                  count: 'text-3xl font-black font-mono text-amber-600',
                  pct: 'text-lg font-black font-mono text-amber-500',
                  bar: 'h-full bg-amber-500 rounded-full transition-all'
              } },
            ].map(({ label, count, classes }) => {
              const pct = historicalStats.total ? Math.round((count / historicalStats.total) * 1000) / 10 : 0;
              return (
                <div key={label} className={classes.card}>
                  <div className="flex items-baseline justify-between mb-2">
                    <span className={classes.count}>{count}</span>
                    <span className={classes.pct}>{pct}%</span>
                  </div>
                  <div className="text-[11px] uppercase font-bold tracking-wider text-slate-500 mb-2">{label}</div>
                  <div className="w-full bg-white rounded-full h-1.5 border border-slate-100 overflow-hidden">
                    <div className={classes.bar} style={{ width: `${pct}%` }} />
                  </div>
                </div>
              );
            })}
          </div>
        )}
        <div className="mt-4 pt-4 border-t border-slate-100 text-xs text-slate-500 font-medium">
          {historicalStats.total} itens avaliados no total
          {historicalStats.naoAvaliado > 0 && ` · ${historicalStats.naoAvaliado} sem avaliação registrada`}
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-5">
        {/* Status do checklist em andamento */}
        <div className="bg-slate-900 rounded-2xl shadow-lg p-6 text-white relative overflow-hidden lg:col-span-1">
          <div className="absolute -right-10 -bottom-10 w-40 h-40 bg-indigo-600/20 rounded-full"></div>
          <h3 className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-4 relative">Auditoria em Andamento</h3>
          <div className="relative">
            <div className="flex items-baseline gap-2 mb-1">
              <span className="text-3xl font-black font-mono">{getProgressPct()}%</span>
              <span className="text-sm text-slate-400 font-medium">concluído</span>
            </div>
            <div className="w-full h-2 bg-slate-700 rounded-full overflow-hidden mb-5">
              <div className="h-full bg-gradient-to-r from-indigo-500 to-indigo-400 rounded-full transition-all" style={{ width: `${getProgressPct()}%` }}></div>
            </div>
            <div className="grid grid-cols-3 gap-3 text-center">
              <div>
                <div className="text-xl font-black text-emerald-400 font-mono">{getCCount()}</div>
                <div className="text-[10px] text-slate-400 uppercase font-bold tracking-wider mt-0.5">Conforme</div>
              </div>
              <div>
                <div className="text-xl font-black text-rose-400 font-mono">{getNcCount()}</div>
                <div className="text-[10px] text-slate-400 uppercase font-bold tracking-wider mt-0.5">Não Conf.</div>
              </div>
              <div>
                <div className="text-xl font-black text-amber-400 font-mono">{getObsCount()}</div>
                <div className="text-[10px] text-slate-400 uppercase font-bold tracking-wider mt-0.5">Observação</div>
              </div>
            </div>
            <button onClick={() => setActiveTab('checklist')} className="mt-5 w-full bg-white/10 hover:bg-white/20 text-white text-sm font-bold py-2.5 rounded-lg transition flex items-center justify-center gap-2">
              Continuar Checklist <ChevronDown size={14} className="-rotate-90" />
            </button>
          </div>
        </div>

        {/* Recents Table */}
        <div className="bg-white rounded-2xl shadow-sm border border-slate-100 overflow-hidden lg:col-span-2">
          <div className="p-5 border-b border-slate-100 flex justify-between items-center">
            <h3 className="text-base font-bold text-slate-800">Últimas Auditorias Concluídas</h3>
            <button onClick={() => setActiveTab('gestao')} className="text-xs text-indigo-600 font-bold hover:underline">Ver Histórico Completo</button>
          </div>
          {savedAudits.length === 0 ? (
            <div className="p-10 text-center text-slate-400 text-sm font-medium">Nenhuma auditoria concluída ainda.</div>
          ) : (
            <table className="w-full text-left border-collapse font-sans text-sm">
              <thead>
                <tr className="bg-slate-50 text-slate-400 text-[11px] uppercase tracking-wider">
                  <th className="p-4 font-bold">Nº RAI</th>
                  <th className="p-4 font-bold">Data</th>
                  <th className="p-4 font-bold">Setor</th>
                  <th className="p-4 font-bold">Auditor</th>
                  <th className="p-4 font-bold text-center">NCs</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {savedAudits.slice(0, 4).map((audit) => (
                  <tr
                    key={audit.id}
                    onClick={() => loadAudit(audit)}
                    className="hover:bg-indigo-50/60 transition-colors cursor-pointer"
                  >
                    <td className="p-4 font-bold text-slate-800 font-mono">{audit.raiNumber}</td>
                    <td className="p-4 text-slate-500">{audit.date}</td>
                    <td className="p-4 font-semibold text-slate-700">{audit.sector}</td>
                    <td className="p-4 text-slate-500">{audit.auditor}</td>
                    <td className="p-4 text-center">
                      <span className={`px-2.5 py-1 rounded-full text-xs font-black ${audit.ncCount > 0 ? 'bg-rose-50 text-rose-600' : 'bg-emerald-50 text-emerald-600'}`}>
                        {audit.ncCount}
                      </span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
      </div>

      {/* Próximas auditorias programadas (vindas do Calendário) */}
      <div className="bg-white rounded-2xl shadow-sm border border-slate-100 overflow-hidden">
        <div className="p-5 border-b border-slate-100 flex justify-between items-center">
          <h3 className="text-base font-bold text-slate-800 flex items-center gap-2">
            <CalendarDays size={18} className="text-indigo-600" /> Próximas Auditorias Programadas
          </h3>
          <button onClick={() => setActiveTab('calendario')} className="text-xs text-indigo-600 font-bold hover:underline">Abrir Calendário</button>
        </div>
        {upcomingEvents.length === 0 ? (
          <div className="p-10 text-center text-slate-400 text-sm font-medium">
            Nenhuma auditoria programada. Use o Calendário para agendar.
          </div>
        ) : (
          <div className="divide-y divide-slate-100">
            {upcomingEvents.slice(0, 4).map((ev) => {
              const evDate = dateKeyToDate(ev.dateKey);
              return (
                <button
                  key={ev.id}
                  onClick={() => { handleSectorChange(ev.sector); setActiveTab('checklist'); }}
                  className="w-full flex items-center gap-4 p-4 hover:bg-indigo-50/40 transition-colors text-left"
                >
                  <div className="w-12 h-12 rounded-xl bg-indigo-50 text-indigo-700 flex flex-col items-center justify-center shrink-0 border border-indigo-100">
                    <span className="text-[10px] font-bold uppercase leading-none">{MONTH_LABELS[evDate.getMonth()].slice(0, 3)}</span>
                    <span className="text-lg font-black leading-none">{evDate.getDate()}</span>
                  </div>
                  <div className="flex-1">
                    <p className="font-bold text-slate-800 text-sm">{ev.sector}</p>
                    <p className="text-xs text-slate-500">{formatTimeRange(ev.time, ev.durationMin)} · {ev.auditor}</p>
                  </div>
                  <ChevronRight size={16} className="text-slate-300" />
                </button>
              );
            })}
          </div>
        )}
      </div>

    </div>
  );

  // ── TELA: Minhas Auditorias em Andamento ────────────────────────────────
  const SETORES_LISTA = sectors.length > 0
    ? sectors.map(s => s.nome)
    : Object.keys(templateMap).length > 0 ? Object.keys(templateMap) : ['COMERCIAL','COMPRAS','PCP','LOGÍSTICA','FISCAL','MANUTENÇÃO','PRODUÇÃO - CALDERARIA','PRODUÇÃO - REVESTIMENTO','PRODUÇÃO - VULCANIZAÇÃO','RECURSOS HUMANOS','INSPEÇÃO DE QUALIDADE','ORÇAMENTO E DESENVOLVIMENTO','PROJETO - DESENHO','ALTA DIREÇÃO','QUALIDADE - SEGURANÇA NO TRABALHO'];

  const [novoSetor, setNovoSetor] = useState('');
  const [showNovaAuditoria, setShowNovaAuditoria] = useState(false);

  const renderCarregandoAuditoria = () => (
    <div className="flex flex-col items-center justify-center py-24 text-slate-400 animate-fade-in">
      <div className="w-10 h-10 border-4 border-indigo-200 border-t-indigo-600 rounded-full animate-spin mb-4"></div>
      <p className="text-sm font-bold">Carregando auditoria…</p>
    </div>
  );

  const renderMinhasAuditorias = () => (
    <div className="flex-1 overflow-y-auto p-6 md:p-8 bg-slate-50 min-h-full">
      <div className="max-w-4xl mx-auto space-y-6">

        {/* Header */}
        <div className="flex items-center justify-between">
          <div>
            <h2 className="text-2xl font-black text-slate-900">Minhas Auditorias</h2>
            <p className="text-slate-500 text-sm mt-1">Cada auditoria é independente — o setor fica travado após criar.</p>
          </div>
          <button
            onClick={() => setShowNovaAuditoria(true)}
            className="flex items-center gap-2 bg-indigo-600 text-white px-5 py-2.5 rounded-xl font-bold hover:bg-indigo-700 transition text-sm shadow-lg shadow-indigo-600/25"
          >
            <Plus size={16} /> Nova Auditoria
          </button>
        </div>

        {/* Modal nova auditoria */}
        {showNovaAuditoria && (
          <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4">
            <div className="bg-white rounded-2xl shadow-2xl w-full max-w-md p-6 space-y-4">
              <h3 className="text-lg font-black text-slate-900">Nova Auditoria</h3>
              <p className="text-sm text-slate-500">Selecione o setor. <strong>Não será possível trocar depois.</strong></p>
              <select
                value={novoSetor}
                onChange={e => setNovoSetor(e.target.value)}
                className="w-full border border-slate-200 rounded-xl p-3 text-sm focus:ring-2 focus:ring-indigo-500 focus:outline-none"
              >
                <option value="">-- Selecione o setor --</option>
                {SETORES_LISTA.map(s => <option key={s} value={s}>{s}</option>)}
              </select>
              <div className="flex gap-3">
                <button
                  onClick={() => { setShowNovaAuditoria(false); setNovoSetor(''); }}
                  className="flex-1 border border-slate-200 text-slate-600 py-2.5 rounded-xl font-bold hover:bg-slate-50 transition text-sm"
                >
                  Cancelar
                </button>
                <button
                  onClick={() => {
                    if (!novoSetor) return;
                    criarNovaAuditoria(novoSetor);
                    setShowNovaAuditoria(false);
                    setNovoSetor('');
                  }}
                  disabled={!novoSetor}
                  className={`flex-1 py-2.5 rounded-xl font-bold transition text-sm ${novoSetor ? 'bg-indigo-600 text-white hover:bg-indigo-700' : 'bg-slate-100 text-slate-400 cursor-not-allowed'}`}
                >
                  Criar Auditoria
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Lista de auditorias em andamento (rascunhos) */}
        <div>
          <h3 className="text-xs font-black text-slate-400 uppercase tracking-wider mb-3">Em andamento</h3>
          {auditorias.length === 0 ? (
            <div className="bg-white rounded-2xl border border-slate-200 p-12 text-center">
              <ClipboardCheck size={40} className="mx-auto text-slate-300 mb-3" />
              <p className="text-slate-500 font-medium">Nenhuma auditoria em andamento.</p>
              <p className="text-slate-400 text-sm mt-1">Clique em "Nova Auditoria" para começar.</p>
            </div>
          ) : (
            <div className="space-y-3">
              {auditorias.map(aud => {
                const respondidos = aud.checklist.filter(i => i.status !== '').length;
                const total       = aud.checklist.length;
                const ncs         = aud.checklist.filter(i => i.status === 'NC').length;
                const pct         = total > 0 ? Math.round((respondidos / total) * 100) : 0;
                const isAtiva     = aud.localId === auditoriaAtivaId;

                return (
                  <div key={aud.localId} className={`bg-white rounded-2xl border-2 p-5 transition ${isAtiva ? 'border-indigo-400 shadow-lg shadow-indigo-100' : 'border-slate-200 hover:border-slate-300'}`}>
                    <div className="flex items-start justify-between gap-4">
                      <div className="flex-1">
                        <div className="flex items-center gap-2 mb-1">
                          <span className="text-xs font-black uppercase tracking-wider px-2 py-0.5 rounded-full bg-amber-100 text-amber-700">Rascunho</span>
                          <span className="text-xs font-bold text-indigo-600 bg-indigo-50 px-2 py-0.5 rounded-full">{aud.raiNumber}</span>
                          {ncs > 0 && <span className="text-xs font-bold bg-rose-100 text-rose-700 px-2 py-0.5 rounded-full">{ncs} NC{ncs > 1 ? 's' : ''}</span>}
                        </div>
                        <h3 className="font-black text-slate-900">{aud.setor}</h3>
                        <div className="mt-2">
                          <div className="flex items-center justify-between text-xs text-slate-500 mb-1">
                            <span>Progresso do checklist</span>
                            <span>{respondidos}/{total} ({pct}%)</span>
                          </div>
                          <div className="w-full bg-slate-100 rounded-full h-1.5">
                            <div className="bg-indigo-500 h-1.5 rounded-full transition-all" style={{ width: `${pct}%` }} />
                          </div>
                        </div>
                      </div>
                      <div className="flex gap-2">
                        <button
                          onClick={() => { setAuditoriaAtivaId(aud.localId); setActiveTab('checklist'); }}
                          className="bg-indigo-600 text-white px-4 py-2 rounded-xl text-xs font-bold hover:bg-indigo-700 transition"
                        >
                          {isAtiva ? 'Continuar' : 'Abrir'}
                        </button>
                        <button
                          onClick={() => requestConfirm({
                            title: 'Descartar rascunho',
                            message: `Deseja descartar a auditoria em andamento do setor ${aud.setor}? Esta ação não pode ser desfeita.`,
                            confirmLabel: 'Sim, descartar',
                            danger: true,
                            onConfirm: () => removerRascunho(aud.localId)
                          })}
                          className="border border-rose-200 text-rose-500 p-2 rounded-xl hover:bg-rose-50 transition"
                        >
                          <Trash2 size={14} />
                        </button>
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>

        {/* Lista de auditorias já concluídas */}
        <div>
          <h3 className="text-xs font-black text-slate-400 uppercase tracking-wider mb-3 mt-6">Concluídas</h3>
          {savedAudits.length === 0 ? (
            <div className="bg-white rounded-2xl border border-slate-200 p-12 text-center">
              <FileText size={40} className="mx-auto text-slate-300 mb-3" />
              <p className="text-slate-500 font-medium">Nenhuma auditoria concluída ainda.</p>
            </div>
          ) : (
            <div className="space-y-3">
              {savedAudits.map(audit => (
                <div key={audit.id} className="bg-white rounded-2xl border-2 border-slate-200 hover:border-slate-300 p-5 transition">
                  <div className="flex items-start justify-between gap-4">
                    <div className="flex-1">
                      <div className="flex items-center gap-2 mb-1">
                        <span className="text-xs font-black uppercase tracking-wider px-2 py-0.5 rounded-full bg-slate-800 text-white">Concluída</span>
                        <span className="text-xs font-bold text-indigo-600 bg-indigo-50 px-2 py-0.5 rounded-full">{audit.raiNumber}</span>
                        <span className={`text-xs font-bold px-2 py-0.5 rounded-full ${audit.ncCount > 0 ? 'bg-rose-100 text-rose-700' : 'bg-emerald-100 text-emerald-700'}`}>
                          {audit.ncCount} NC{audit.ncCount !== 1 ? 's' : ''}
                        </span>
                      </div>
                      <h3 className="font-black text-slate-900">{audit.sector}</h3>
                      <p className="text-xs text-slate-500 mt-1">{audit.date} · Auditor: {audit.auditor} · {audit.branch}</p>
                    </div>
                    <button
                      onClick={() => loadAudit(audit)}
                      disabled={loadingAuditId === audit.id}
                      className="bg-white text-indigo-600 border border-indigo-200 hover:bg-indigo-600 hover:text-white px-4 py-2 rounded-xl text-xs font-bold transition-all shadow-sm disabled:opacity-50 disabled:cursor-wait shrink-0"
                    >
                      {loadingAuditId === audit.id ? '...' : 'VER'}
                    </button>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  );

  const renderChecklist = () => (
    <div className="space-y-6 animate-fade-in">
      <div className="flex justify-between items-center mb-4">
        <div>
          <h2 className="text-2xl font-black text-slate-900 flex items-center gap-3">
            Lista de Verificação
            <span className={`text-[11px] font-black uppercase tracking-wider px-3 py-1 rounded-full ${
              isChecklistLocked ? 'bg-slate-800 text-white' : 'bg-emerald-100 text-emerald-700'
            }`}>
              {isChecklistLocked ? 'Fechado' : 'Em Andamento'}
            </span>
          </h2>
          <p className="text-slate-500 text-sm mt-1">
            {getAnsweredCount()} de {checklist.length} itens avaliados · {selectedSector}
            {isChecklistLocked && checklistClosedAt && (
              <> · Fechado em {new Date(checklistClosedAt).toLocaleDateString('pt-BR')} às {new Date(checklistClosedAt).toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' })}</>
            )}
          </p>
        </div>
        <div className="flex gap-3 flex-wrap">
          <button onClick={exportChecklistExcel} className="flex items-center gap-2 bg-white border border-slate-200 text-emerald-700 px-4 py-2.5 rounded-xl font-bold hover:bg-emerald-50 transition text-sm shadow-sm">
            <FileSpreadsheet size={16} /> Exportar Excel
          </button>
          <button onClick={printDocument} className="flex items-center gap-2 bg-white border border-slate-200 text-slate-700 px-4 py-2.5 rounded-xl font-bold hover:bg-slate-50 transition text-sm shadow-sm">
            <Printer size={16} /> Imprimir / PDF
          </button>
          {!isChecklistLocked && (
            <button
              onClick={saveRascunho}
              className={`flex items-center gap-2 px-4 py-2.5 rounded-xl font-bold transition text-sm border ${
                autoSaveStatus === 'saved'
                  ? 'bg-emerald-50 border-emerald-200 text-emerald-700 hover:bg-emerald-100'
                  : 'bg-amber-50 border-amber-300 text-amber-700 hover:bg-amber-100'
              }`}
            >
              <Save size={16} />
              {autoSaveStatus === 'saved' ? 'Rascunho salvo ✓' : 'Salvar Rascunho'}
            </button>
          )}
          {isChecklistLocked ? (
            <button onClick={requestReopenChecklist} className="flex items-center gap-2 bg-amber-500 text-white px-4 py-2.5 rounded-xl font-bold hover:bg-amber-600 transition text-sm shadow-lg shadow-amber-500/25">
              <FolderTree size={16} /> Reabrir Checklist
            </button>
          ) : (
            <button onClick={requestCloseChecklist} className="flex items-center gap-2 bg-slate-800 text-white px-4 py-2.5 rounded-xl font-bold hover:bg-slate-900 transition text-sm shadow-lg shadow-slate-800/25">
              <CheckCircle size={16} /> Fechar Checklist
            </button>
          )}
        </div>
      </div>

      {isChecklistLocked && (
        <div className="bg-slate-50 border border-slate-200 text-slate-600 rounded-xl p-4 flex items-start gap-3 text-sm max-w-[1000px] mx-auto print:hidden">
          <ShieldCheck size={18} className="mt-0.5 shrink-0 text-slate-400" />
          <div>
            Este checklist está <span className="font-bold">fechado</span> e não pode ser editado. Use "Reabrir Checklist" se precisar corrigir algo —
            a reabertura fica registrada abaixo para rastreabilidade.
          </div>
        </div>
      )}

      {reopenHistory.length > 0 && (
        <div className="bg-amber-50 border border-amber-200 text-amber-800 rounded-xl p-4 text-sm max-w-[1000px] mx-auto print:hidden">
          <p className="font-bold mb-1.5 flex items-center gap-2"><Clock size={14} /> Histórico de reaberturas</p>
          <ul className="space-y-1 ml-1">
            {reopenHistory.map((r, i) => (
              <li key={i} className="text-xs">
                Reaberto em {new Date(r.date).toLocaleDateString('pt-BR')} às {new Date(r.date).toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' })} por <span className="font-bold">{r.auditor}</span>
              </li>
            ))}
          </ul>
        </div>
      )}

      {usingGenericChecklist && !isChecklistLocked && (
        <div className="bg-amber-50 border border-amber-200 text-amber-800 rounded-xl p-4 flex items-start gap-3 text-sm max-w-[1000px] mx-auto print:hidden">
          <AlertCircle size={18} className="mt-0.5 shrink-0" />
          <div>
            <span className="font-bold">{selectedSector}</span> ainda não tem um roteiro de checklist específico cadastrado.
            Está sendo usado o roteiro genérico do SGQ — edite as perguntas abaixo ou adicione novas conforme a realidade do setor.
          </div>
        </div>
      )}

      {ncItemsWithoutRnc.length > 0 && (
        <div className="bg-rose-50 border border-rose-200 text-rose-800 rounded-xl p-4 flex items-start justify-between gap-3 text-sm max-w-[1000px] mx-auto print:hidden">
          <div className="flex items-start gap-3">
            <AlertTriangle size={18} className="mt-0.5 shrink-0" />
            <div>
              <span className="font-bold">{ncItemsWithoutRnc.length} item(ns)</span> marcado(s) como Não Conforme ainda sem RNC aberta.
            </div>
          </div>
          <button
            onClick={() => createRnc(ncItemsWithoutRnc[0])}
            className="shrink-0 bg-rose-600 text-white text-xs font-bold px-3 py-1.5 rounded-lg hover:bg-rose-700 transition flex items-center gap-1.5 whitespace-nowrap"
          >
            <Link2 size={13} /> Abrir RNC para este item
          </button>
        </div>
      )}

      <div className={`bg-white shadow-xl border-2 border-black font-sans text-sm max-w-[1000px] mx-auto print:shadow-none print:border-0 print:w-full relative ${isChecklistLocked ? 'opacity-90' : ''}`}>
        <div className="flex border-b-2 border-black">
          <div className="w-1/4 p-2 flex flex-col items-center justify-center border-r-2 border-black bg-white">
            <KalenLogo variant="full" />
          </div>
          <div className="w-1/2 p-4 flex items-center justify-center border-r-2 border-black">
            <h1 className="text-xl font-bold text-center uppercase leading-tight">
              KdB147 - LISTA DE VERIFICAÇÃO DE<br />AUDITORIA INTERNA
            </h1>
          </div>
          <div className="w-1/4 flex flex-col">
            <div className="flex flex-1 border-b-2 border-black">
              <div className="w-1/2 border-r-2 border-black flex items-center justify-center font-bold bg-gray-200 text-xs">SETOR</div>
              <div className="w-1/2 flex items-center justify-center">
                <select
                  value={selectedSector}
                  onChange={(e) => handleSectorChange(e.target.value)}
                  disabled={isChecklistLocked}
                  className="w-full h-full text-center text-xs font-bold bg-transparent focus:outline-none cursor-pointer disabled:cursor-not-allowed disabled:opacity-60"
                >
                  {SECTORS.map((s) => <option key={s} value={s}>{s}</option>)}
                </select>
              </div>
            </div>
            <div className="flex flex-1">
              <div className="w-1/2 border-r-2 border-black flex items-center justify-center font-bold bg-gray-100 text-xs">DATA</div>
              <div className="w-1/2 flex items-center justify-center text-xs font-bold">{formattedDate}</div>
            </div>
          </div>
        </div>

        {selectedSectorRequisitos && (
          <div className="flex border-b-2 border-black">
            <div className="w-[15%] py-2 px-2 border-r-2 border-black text-center font-bold bg-gray-100 text-xs">Requisitos ISO:</div>
            <div className="w-[85%] py-2 px-4 font-bold text-xs text-blue-900">{selectedSectorRequisitos}</div>
          </div>
        )}

        <div className="flex border-b-2 border-black">
          <div className="w-[15%] py-2 px-2 border-r-2 border-black text-center font-bold bg-gray-100 text-xs">Documentos:</div>
          <div className="w-[85%] py-2 px-3 text-xs text-gray-600">Norma NBR ISO 9001:2015 / Manual SGQ / Procedimentos / Mapa de Processo / Formulários</div>
        </div>

        <div className="flex border-b-2 border-black">
          <div className="w-[15%] py-2 px-2 border-r-2 border-black text-center font-bold">Legenda:</div>
          <div className="w-[85%] py-2 px-4 font-bold flex gap-6">
            <span className="flex items-center gap-1"><div className="w-3 h-3 bg-green-500 rounded-sm"></div> C - Conforme</span>
            <span className="flex items-center gap-1"><div className="w-3 h-3 bg-red-500 rounded-sm"></div> NC - Não Conforme</span>
            <span className="flex items-center gap-1"><div className="w-3 h-3 bg-blue-500 rounded-sm"></div> Obs - Observação</span>
          </div>
        </div>

        <table className="w-full text-left border-collapse">
          <thead>
            <tr className="bg-gray-200 text-black font-bold text-sm border-b-2 border-black">
              <th className="p-3 w-12 text-center border-r-2 border-black uppercase">Nº</th>
              <th className="p-3 w-1/2 border-r-2 border-black text-center uppercase">Questionamento Operacional</th>
              <th className="p-3 border-r-2 border-black text-center uppercase">Evidências / Comentários</th>
              <th className="p-3 w-32 text-center uppercase">Avaliação</th>
              {!isChecklistLocked && <th className="p-3 w-10 text-center uppercase print:hidden"></th>}
            </tr>
          </thead>
          <tbody className="divide-y-2 divide-black">
            {checklist.map((item, idx) => (
              <tr key={item.id} className="hover:bg-blue-50 transition-colors group">
                <td className="p-3 text-center text-black font-bold border-r-2 border-black align-top bg-gray-50 group-hover:bg-blue-100">{idx + 1}</td>
                <td className="p-3 text-sm text-black border-r-2 border-black align-top font-medium">
                  <textarea
                    rows="2"
                    value={item.question}
                    onChange={(e) => handleChecklistChange(item.id, 'question', e.target.value)}
                    placeholder="Digite a pergunta do roteiro de auditoria..."
                    readOnly={isChecklistLocked}
                    className="w-full text-sm p-0 bg-transparent border-0 focus:ring-0 focus:outline-none resize-y min-h-[40px] font-medium disabled:cursor-default"
                  ></textarea>
                </td>
                <td className="p-3 border-r-2 border-black align-top">
                  <textarea
                    rows="2" value={item.comments} onChange={(e) => handleChecklistChange(item.id, 'comments', e.target.value)}
                    placeholder="Registrar evidência objetiva..."
                    readOnly={isChecklistLocked}
                    className="w-full text-sm p-2 bg-transparent border-0 focus:ring-0 focus:outline-none resize-y min-h-[60px] italic text-gray-700"
                  ></textarea>
                </td>
                <td className="p-3 align-top bg-white">
                  <div className="flex flex-col gap-1">
                    {['C', 'NC', 'Obs'].map((statusOption) => (
                      <button
                        key={statusOption}
                        onClick={() => handleChecklistChange(item.id, 'status', statusOption)}
                        disabled={isChecklistLocked}
                        className={`px-2 py-1 rounded-sm text-xs font-bold border transition-all print:border-black print:text-black ${
                          item.status === statusOption
                            ? statusOption === 'C' ? 'bg-green-600 text-white border-green-700 print:bg-gray-300'
                              : statusOption === 'NC' ? 'bg-red-600 text-white border-red-700 print:bg-gray-300'
                              : 'bg-blue-600 text-white border-blue-700 print:bg-gray-300'
                            : 'bg-gray-100 text-gray-500 border-gray-300 hover:bg-gray-200'
                        } ${isChecklistLocked ? 'cursor-not-allowed opacity-80' : ''}`}
                      >
                        {statusOption}
                      </button>
                    ))}
                  </div>
                </td>
                {!isChecklistLocked && (
                  <td className="p-2 align-top text-center print:hidden">
                    <button
                      onClick={() => requestConfirm({
                        title: 'Remover item do checklist',
                        message: 'Tem certeza que deseja remover este item da lista de verificação?',
                        confirmLabel: 'Sim, remover',
                        danger: true,
                        onConfirm: () => handleRemoveChecklistItem(item.id)
                      })}
                      aria-label="Remover item"
                      className="text-slate-300 hover:text-rose-600 transition"
                    >
                      <Trash2 size={15} />
                    </button>
                  </td>
                )}
              </tr>
            ))}
          </tbody>
        </table>

        {!isChecklistLocked && (
          <div className="p-3 bg-gray-50 border-t-2 border-black print:hidden">
            <button
              onClick={handleAddChecklistItem}
              className="flex items-center gap-2 text-indigo-600 font-bold text-sm hover:text-indigo-800 transition"
            >
              <ListPlus size={16} /> Adicionar item ao checklist
            </button>
          </div>
        )}
      </div>
    </div>
  );


  const renderReport = () => {
    const saveBlockers = getSaveBlockers();
    return (
    <div className="space-y-6 animate-fade-in">
      <div className="flex justify-between items-center mb-4">
        <div>
          <h2 className="text-2xl font-black text-slate-900">Emissão de Relatório Final</h2>
          <p className="text-slate-500 text-sm mt-1 flex items-center gap-2">
            RAI Nº <span className="font-mono font-bold text-indigo-600">{report.raiNumber}</span>
            <span className={`text-[10px] font-black uppercase tracking-wider px-2.5 py-0.5 rounded-full ${
              isChecklistLocked ? 'bg-slate-800 text-white' : 'bg-emerald-100 text-emerald-700'
            }`}>
              Checklist {isChecklistLocked ? 'Fechado' : 'Em Andamento'}
            </span>
          </p>
        </div>
        <div className="flex gap-3 flex-wrap">
          <button onClick={exportReportExcel} className="flex items-center gap-2 bg-white border border-slate-200 text-emerald-700 px-4 py-2.5 rounded-xl font-bold hover:bg-emerald-50 transition text-sm shadow-sm">
            <FileSpreadsheet size={16} /> Exportar Excel
          </button>
          <button onClick={printDocument} className="flex items-center gap-2 bg-white border border-slate-200 text-slate-700 px-4 py-2.5 rounded-xl font-bold hover:bg-slate-50 transition text-sm shadow-sm">
            <Printer size={16} /> Imprimir / PDF
          </button>
          {auditoriaAtivaId && (
            <button
              onClick={saveRascunho}
              className={`flex items-center gap-2 px-4 py-2.5 rounded-xl font-bold transition text-sm border ${
                autoSaveStatus === 'saved'
                  ? 'bg-emerald-50 border-emerald-200 text-emerald-700 hover:bg-emerald-100'
                  : 'bg-amber-50 border-amber-300 text-amber-700 hover:bg-amber-100'
              }`}
            >
              <Save size={16} />
              {autoSaveStatus === 'saved' ? 'Rascunho salvo ✓' : 'Salvar Rascunho'}
            </button>
          )}
          <button
            onClick={() => {
              if (saveBlockers.length > 0) {
                notify('Auditoria incompleta', saveBlockers.join(' '));
                return;
              }
              requestConfirm({
                title: editingAuditId ? 'Salvar alterações na auditoria' : 'Confirmar salvamento da auditoria',
                message: editingAuditId
                  ? `Tem certeza que deseja salvar as alterações na auditoria RAI Nº ${report.raiNumber}? O histórico será atualizado.`
                  : `Tem certeza que deseja finalizar e salvar a auditoria RAI Nº ${report.raiNumber} no histórico?`,
                confirmLabel: editingAuditId ? 'Sim, salvar alterações' : 'Sim, salvar auditoria',
                danger: false,
                onConfirm: handleSaveAudit
              });
            }}
            className={`flex items-center gap-2 px-4 py-2.5 rounded-xl font-bold transition text-sm shadow-lg ${
              saveBlockers.length > 0
                ? 'bg-slate-200 text-slate-400 cursor-not-allowed shadow-none'
                : editingAuditId
                  ? 'bg-amber-600 text-white hover:bg-amber-700 shadow-amber-600/25'
                  : 'bg-indigo-600 text-white hover:bg-indigo-700 shadow-indigo-600/25'
            }`}
          >
            <Save size={16} /> {editingAuditId ? 'Salvar Alterações' : 'Salvar no Histórico'}
          </button>
        </div>
      </div>

      {/* Banner de modo de edição */}
      {editingAuditId && (
        <div className="bg-amber-50 border border-amber-300 text-amber-800 rounded-xl p-4 flex items-center gap-3 text-sm max-w-[1000px] mx-auto print:hidden">
          <Edit3 size={18} className="shrink-0" />
          <div className="flex-1">
            <span className="font-bold">Modo de edição</span> — você está editando uma auditoria já salva no histórico.
            Suas alterações serão aplicadas ao registro existente ao salvar.
          </div>
          <button
            onClick={() => { setEditingAuditId(null); setActiveTab('gestao'); }}
            className="text-amber-600 hover:text-amber-800 font-bold text-xs border border-amber-300 hover:border-amber-500 px-3 py-1.5 rounded-lg transition"
          >
            Cancelar edição
          </button>
        </div>
      )}


      {saveBlockers.length > 0 && (
        <div className="bg-amber-50 border border-amber-200 text-amber-800 rounded-xl p-4 flex items-start gap-3 text-sm max-w-[1000px] mx-auto print:hidden">
          <AlertCircle size={18} className="mt-0.5 shrink-0" />
          <div>
            <span className="font-bold">Pendências antes de salvar:</span>
            <ul className="list-disc ml-5 mt-1 space-y-0.5">
              {saveBlockers.map((b, i) => <li key={i}>{b}</li>)}
            </ul>
          </div>
        </div>
      )}

      <div className="bg-white shadow-xl border-2 border-black font-sans text-sm max-w-[1000px] mx-auto print:shadow-none print:border-0 print:w-full">
        <div className="flex border-b-2 border-black">
          <div className="w-1/4 p-2 flex flex-col items-center justify-center border-r-2 border-black bg-white">
            <KalenLogo variant="full" />
          </div>
          <div className="w-1/2 p-4 flex items-center justify-center border-r-2 border-black bg-gray-50">
            <h1 className="text-2xl font-black text-center uppercase tracking-wide text-gray-800">
              Relatório de Auditoria Interna
            </h1>
          </div>
          <div className="w-1/4 flex flex-col">
            <div className="flex flex-1 border-b-2 border-black">
              <div className="w-1/2 border-r-2 border-black flex items-center justify-center font-bold bg-gray-200">RAI Nº</div>
              <div className="w-1/2 flex items-center justify-center font-bold text-red-700 text-lg">{report.raiNumber}</div>
            </div>
            <div className="flex flex-1">
              <div className="w-full flex items-center justify-center bg-gray-100 text-xs font-bold uppercase tracking-widest text-gray-500">Documento Controlado</div>
            </div>
          </div>
        </div>

        <div className="flex border-b-2 border-black">
          <div className="w-[15%] p-2 border-r-2 border-black font-bold bg-gray-100 flex items-center">Auditoria Inicial:</div>
          <div className="w-[20%] p-2 border-r-2 border-black flex items-center justify-center bg-white"><input type="text" className="w-full text-center focus:outline-none bg-transparent" defaultValue={formattedDate} readOnly /></div>
          <div className="w-[15%] p-2 border-r-2 border-black font-bold bg-gray-100 flex items-center">Data de Emissão:</div>
          <div className="w-[15%] p-2 border-r-2 border-black flex items-center justify-center bg-white"><input type="text" className="w-full text-center focus:outline-none bg-transparent" value={report.date} readOnly /></div>
          <div className="w-[10%] p-2 border-r-2 border-black font-bold bg-gray-100 flex items-center">Unidade:</div>
          <div className="w-[25%] p-2 flex items-center justify-center bg-white">
            <select value={selectedBranch} onChange={(e) => setSelectedBranch(e.target.value)} className="w-full text-center focus:outline-none font-bold bg-transparent cursor-pointer">
              {BRANCHES.map((b) => <option key={b} value={b}>{b}</option>)}
            </select>
          </div>
        </div>

        <div className="flex border-b-2 border-black bg-white">
          <div className="w-[15%] p-2 border-r-2 border-black font-bold bg-gray-100 flex items-center">Área:</div>
          <div className="w-[85%] p-2">
            <select
              value={selectedSector}
              onChange={(e) => handleSectorChange(e.target.value)}
              disabled={isChecklistLocked}
              className="w-full focus:outline-none font-bold text-blue-900 uppercase bg-transparent cursor-pointer disabled:cursor-not-allowed disabled:opacity-60"
            >
              {SECTORS.map((s) => <option key={s} value={s}>{s}</option>)}
            </select>
          </div>
        </div>
        <div className="flex border-b-2 border-black bg-white">
          <div className="w-[15%] p-2 border-r-2 border-black font-bold bg-gray-100 flex items-center">Processo:</div>
          <div className="w-[85%] p-2"><input type="text" className="w-full focus:outline-none font-bold text-blue-900 uppercase bg-transparent" value={selectedSector} readOnly /></div>
        </div>

        <div className="flex border-b-2 border-black bg-white">
          <div className="w-[15%] p-2 border-r-2 border-black font-bold bg-gray-100 flex items-center">Metodologia:</div>
          <div className="w-[85%] p-2"><input type="text" className="w-full focus:outline-none bg-transparent" defaultValue="Entrevista, Amostragem e Análise de Documentos do SGQ" /></div>
        </div>
        <div className="flex border-b-2 border-black bg-white">
          <div className="w-[15%] p-2 border-r-2 border-black font-bold bg-gray-100 flex items-center">Objetivo:</div>
          <div className="w-[85%] p-2"><textarea className="w-full focus:outline-none resize-none bg-transparent" rows="2" defaultValue="Verificar a conformidade e eficácia do Processo dentro do Sistema da Gestão de Qualidade Kalenborn, garantindo alinhamento à ISO 9001:2015."></textarea></div>
        </div>

        <div className="flex border-b-2 border-black bg-white">
          <div className="w-[20%] p-2 border-r-2 border-black font-bold bg-gray-100 flex items-center">Auditor Líder:</div>
          <div className="w-[80%] p-2">
            <select value={report.auditor} onChange={(e) => setReport({ ...report, auditor: e.target.value })} className="w-full focus:outline-none bg-transparent font-bold cursor-pointer">
              {auditoresDb.map((a) => <option key={a.id} value={a.nome}>{a.nome}</option>)}
            </select>
          </div>
        </div>
        <div className="flex border-b-2 border-black bg-white">
          <div className="w-[20%] p-2 border-r-2 border-black font-bold bg-gray-100 flex items-center">Equipe / Convidados:</div>
          <div className="w-[80%] p-2"><input type="text" className="w-full focus:outline-none bg-transparent" defaultValue="Comitê de Qualidade da Unidade" /></div>
        </div>

        <div className="bg-gray-300 font-bold p-2 border-b-2 border-black text-center uppercase tracking-wider">Pessoas Auditadas (Entrevistados)</div>
        <div className="flex border-b-2 border-black text-center bg-gray-50 font-bold">
          <div className="w-1/2 p-2 border-r-2 border-black">Nome do Colaborador</div>
          <div className="w-1/2 p-2">Cargo / Função</div>
        </div>
        <div className="flex border-b-2 border-black bg-white hover:bg-gray-50">
          <div className="w-1/2 p-2 border-r-2 border-black">
            <input
              type="text"
              className="w-full text-center focus:outline-none bg-transparent"
              placeholder="Nome do colaborador entrevistado"
              value={report.auditee}
              onChange={(e) => setReport({ ...report, auditee: e.target.value })}
            />
          </div>
          <div className="w-1/2 p-2"><input type="text" className="w-full text-center focus:outline-none bg-transparent" defaultValue={`Assistente de ${selectedSector}`} /></div>
        </div>

        <div className="bg-gray-300 font-bold p-2 border-b-2 border-black text-center uppercase tracking-wider">Pontos Fortes / Positivos Identificados</div>
        <div className="border-b-2 border-black bg-white p-4">
          <textarea className="w-full focus:outline-none resize-y min-h-[80px] bg-transparent" placeholder="Ex.: 1 - Cordialidade do auditado..." value={report.positivePoints} onChange={(e) => setReport({ ...report, positivePoints: e.target.value })}></textarea>
        </div>

        <div className="bg-gray-300 font-bold p-2 border-b-2 border-black text-center uppercase tracking-wider">Observações Pertinentes (Geradas pelo Checklist)</div>
        <div className="border-b-2 border-black bg-white p-4">
          <textarea className="w-full focus:outline-none resize-y min-h-[80px] bg-transparent" value={report.observations} onChange={(e) => setReport({ ...report, observations: e.target.value })}></textarea>
        </div>

        <div className="bg-gray-300 font-bold p-2 border-b-2 border-black text-center uppercase tracking-wider">Oportunidades e Sugestões de Melhoria</div>
        <div className="border-b-2 border-black bg-white p-4">
          <textarea className="w-full focus:outline-none resize-y min-h-[80px] bg-transparent" placeholder="Ex.: 1 - Definir qual setor será responsável por..." value={report.improvements} onChange={(e) => setReport({ ...report, improvements: e.target.value })}></textarea>
        </div>

        <div className="flex border-b-2 border-black bg-yellow-50">
          <div className="w-3/4 p-3 border-r-2 border-black flex items-center font-bold text-red-800 uppercase">Quantidade de RNCs (Relatórios de Não Conformidade) exigidas para este processo:</div>
          <div className="w-1/4 p-3 flex items-center justify-center font-black text-2xl text-red-600 bg-white">
            {getNcCount()}
          </div>
        </div>

        <div className="bg-gray-300 font-bold p-2 border-b-2 border-black text-center uppercase tracking-wider">Conclusão Oficial da Auditoria</div>
        <div className="border-b-2 border-black bg-white p-4">
          <textarea className="w-full text-center focus:outline-none resize-y min-h-[80px] font-medium bg-transparent" placeholder="Ex.: O processo atende parcialmente aos requisitos do SGQ." value={report.conclusion} onChange={(e) => setReport({ ...report, conclusion: e.target.value })}></textarea>
        </div>

        <div className="flex text-center bg-gray-50">
          <div className="w-1/2 p-6 border-r-2 border-black flex flex-col items-center justify-end">
            <div className="w-3/4 border-b border-black mb-2"></div>
            <div className="font-bold text-gray-800">Assinatura do Auditor Líder</div>
            <div className="text-xs text-gray-500 mt-1">{report.auditor}</div>
          </div>
          <div className="w-1/2 p-6 flex flex-col items-center justify-end">
            <div className="w-3/4 border-b border-black mb-2"></div>
            <div className="font-bold text-gray-800">Responsável SGQ Unidade</div>
            <div className="text-xs text-gray-500 mt-1">Coordenação de Qualidade — {selectedBranch}</div>
          </div>
        </div>
      </div>
    </div>
  );
  };

  const renderRNC = () => {
    const visibleRncs = rncStatusFilter === 'Todas' ? rncs : rncs.filter((r) => r.status === rncStatusFilter);
    return (
    <div className="space-y-6 animate-fade-in">
      <div className="flex justify-between items-center mb-4">
        <div>
          <h2 className="text-2xl font-black text-slate-900">Tratativa de Não Conformidades (RNC)</h2>
          <p className="text-slate-500 text-sm mt-1">{rncs.filter((r) => r.status === 'Aberta').length} aberta(s) · {rncs.length} no total</p>
        </div>
        <div className="flex gap-3">
          <button onClick={exportRncExcel} className="flex items-center gap-2 bg-white border border-slate-200 text-emerald-700 px-4 py-2.5 rounded-xl font-bold hover:bg-emerald-50 transition text-sm shadow-sm">
            <FileSpreadsheet size={16} /> Exportar Excel
          </button>
          <button onClick={printDocument} className="flex items-center gap-2 bg-white border border-slate-200 text-slate-700 px-4 py-2.5 rounded-xl font-bold hover:bg-slate-50 transition text-sm shadow-sm">
            <Printer size={16} /> Imprimir / PDF
          </button>
          <button
            onClick={() => createRnc(null)}
            className="flex items-center gap-2 bg-rose-600 text-white px-4 py-2.5 rounded-xl font-bold hover:bg-rose-700 transition text-sm shadow-lg shadow-rose-600/25"
          >
            <Plus size={16} /> Gerar Nova RNC
          </button>
        </div>
      </div>

      <div className="flex gap-2 print:hidden">
        {['Todas', 'Aberta', 'Fechada'].map((opt) => (
          <button
            key={opt}
            onClick={() => setRncStatusFilter(opt)}
            className={`px-4 py-1.5 rounded-full text-xs font-bold border transition ${
              rncStatusFilter === opt
                ? 'bg-indigo-600 text-white border-indigo-600 shadow-sm'
                : 'bg-white text-slate-500 border-slate-200 hover:bg-slate-50'
            }`}
          >
            {opt} {opt !== 'Todas' && `(${rncs.filter((r) => r.status === opt).length})`}
          </button>
        ))}
      </div>

      {rncs.length === 0 ? (
        <div className="bg-white border border-slate-200 rounded-2xl p-16 text-center text-slate-500 shadow-sm font-medium flex flex-col items-center justify-center">
          <AlertTriangle size={48} className="text-slate-300 mb-4" />
          Nenhuma RNC registrada. Use "Gerar Nova RNC" para abrir um novo relatório de ação corretiva.
        </div>
      ) : visibleRncs.length === 0 ? (
        <div className="bg-white border border-slate-200 rounded-2xl p-16 text-center text-slate-500 shadow-sm font-medium flex flex-col items-center justify-center">
          <ShieldCheck size={48} className="text-slate-300 mb-4" />
          Nenhuma RNC com status "{rncStatusFilter}".
        </div>
      ) : visibleRncs.map((rnc) => (
        <div key={rnc.id} className="bg-white border-2 border-black font-sans text-sm relative mb-8 shadow-xl max-w-[1000px] mx-auto print:shadow-none print:border-0 print:w-full">
          <div className="absolute -top-4 -right-4 print:hidden">
            <span className={`px-4 py-1 text-xs font-black uppercase tracking-wider rounded-full shadow-lg border-2 ${rnc.status === 'Aberta' ? 'bg-amber-400 border-amber-600 text-amber-900' : 'bg-emerald-500 border-emerald-700 text-white'}`}>
              Status: {rnc.status}
            </span>
          </div>

          <div className="flex border-b-2 border-black">
            <div className="w-1/4 p-2 flex flex-col items-center justify-center border-r-2 border-black bg-white">
              <KalenLogo variant="full" />
            </div>
            <div className="w-1/2 p-4 flex items-center justify-center border-r-2 border-black bg-gray-50">
              <h1 className="text-xl font-black text-center uppercase tracking-wide text-gray-800">RELATÓRIO DE AÇÕES CORRETIVAS</h1>
            </div>
            <div className="w-1/4 flex flex-col">
              <div className="flex flex-1 border-b-2 border-black">
                <div className="w-1/2 border-r-2 border-black flex items-center justify-center font-bold bg-gray-200 text-xs">Nº R.A.C</div>
                <div className="w-1/2 flex items-center justify-center font-black text-xl text-red-700 bg-red-50">{rnc.id}</div>
              </div>
              <div className="flex flex-1">
                <div className="w-1/2 border-r-2 border-black flex items-center justify-center font-bold bg-gray-100 text-xs">DATA EMISSÃO:</div>
                <div className="w-1/2 flex items-center justify-center font-bold">{rnc.date}</div>
              </div>
            </div>
          </div>

          {rnc.sourceChecklistItem && (
            <div className="bg-indigo-50 border-b-2 border-black p-2.5 px-3 flex items-center gap-2 text-xs text-indigo-800 print:hidden">
              <Link2 size={14} className="shrink-0" />
              <span>
                Originada do checklist da auditoria <span className="font-bold">RAI {rnc.sourceRaiNumber}</span> — item:
                <span className="font-semibold"> "{rnc.sourceChecklistItem.question}"</span>
              </span>
            </div>
          )}

          <div className="bg-gray-300 font-bold p-1 px-3 border-b-2 border-black uppercase tracking-wider">Origem da Ação</div>
          <div className="flex border-b-2 border-black bg-white p-3 justify-around text-xs font-bold text-gray-700">
            {['AUDITORIA INTERNA', 'AUDITORIA EXTERNA', 'RECLAMAÇÃO DE CLIENTE', 'PRODUTO / PROCESSO'].map((opt) => (
              <label key={opt} className="flex items-center gap-2 cursor-pointer">
                <input
                  type="checkbox"
                  checked={rnc.origin === opt}
                  onChange={() => handleRncChange(rnc.id, 'origin', opt)}
                />
                {opt}
              </label>
            ))}
          </div>

          <div className="bg-gray-300 font-bold p-1 px-3 border-b-2 border-black uppercase tracking-wider">Dados de Origem</div>
          <div className="flex border-b-2 border-black bg-white">
            <div className="w-[15%] font-bold p-2 border-r-2 border-black flex items-center bg-gray-50">Processo:</div>
            <div className="w-[45%] p-2 border-r-2 border-black"><input type="text" className="w-full focus:outline-none uppercase font-bold text-blue-900" value={rnc.process} onChange={(e) => handleRncChange(rnc.id, 'process', e.target.value)} /></div>
            <div className="w-[15%] font-bold p-2 border-r-2 border-black flex items-center bg-gray-50">Unidade:</div>
            <div className="w-[25%] p-2"><input type="text" className="w-full focus:outline-none font-bold" value={selectedBranch} readOnly /></div>
          </div>

          <div className="bg-gray-300 font-bold p-2 border-b-2 border-black text-center uppercase tracking-wider">Descrição Detalhada da Não Conformidade</div>
          <div className="border-b-2 border-black bg-white p-4">
            <textarea className="w-full focus:outline-none resize-y min-h-[80px] text-red-900 font-medium bg-red-50 p-2 rounded" value={rnc.description} onChange={(e) => handleRncChange(rnc.id, 'description', e.target.value)}></textarea>
          </div>

          <div className="bg-gray-300 font-bold p-2 border-b-2 border-black text-center uppercase tracking-wider">Ação de Correção Imediata (Disposição)</div>
          <div className="flex border-b-2 border-black font-bold text-center bg-gray-100 text-xs uppercase">
            <div className="w-[60%] border-r-2 border-black p-2">Descrição da Ação Imediata</div>
            <div className="w-[20%] border-r-2 border-black p-2">Responsável</div>
            <div className="w-[20%] p-2">Data Limite</div>
          </div>
          <div className="flex border-b-2 border-black bg-white">
            <div className="w-[60%] border-r-2 border-black"><textarea className="w-full h-full p-3 focus:outline-none resize-none" value={rnc.correction} onChange={(e) => handleRncChange(rnc.id, 'correction', e.target.value)}></textarea></div>
            <div className="w-[20%] border-r-2 border-black"><input type="text" className="w-full h-full p-3 focus:outline-none text-center" value={rnc.correctionResp} onChange={(e) => handleRncChange(rnc.id, 'correctionResp', e.target.value)} /></div>
            <div className="w-[20%]"><input type="date" className="w-full h-full p-3 focus:outline-none text-center" value={rnc.correctionDate} onChange={(e) => handleRncChange(rnc.id, 'correctionDate', e.target.value)} /></div>
          </div>

          <div className="bg-gray-300 font-bold p-2 border-b-2 border-black text-center uppercase tracking-wider text-red-900">Análise de Causa(s) Raiz</div>
          <div className="flex border-b-2 border-black font-bold text-center bg-gray-100 text-xs uppercase">
            <div className="w-[60%] border-r-2 border-black p-2">Metodologia / Descrição das Causas Encontradas</div>
            <div className="w-[20%] border-r-2 border-black p-2">Líder da Análise</div>
            <div className="w-[20%] p-2">Data Conclusão</div>
          </div>
          <div className="flex border-b-2 border-black bg-white">
            <div className="w-[60%] border-r-2 border-black"><textarea className="w-full h-full p-3 focus:outline-none resize-none bg-red-50/30" value={rnc.rootCause} onChange={(e) => handleRncChange(rnc.id, 'rootCause', e.target.value)}></textarea></div>
            <div className="w-[20%] border-r-2 border-black"><input type="text" className="w-full h-full p-3 focus:outline-none text-center" value={rnc.rootCauseResp} onChange={(e) => handleRncChange(rnc.id, 'rootCauseResp', e.target.value)} /></div>
            <div className="w-[20%]"><input type="date" className="w-full h-full p-3 focus:outline-none text-center" value={rnc.rootCauseDate} onChange={(e) => handleRncChange(rnc.id, 'rootCauseDate', e.target.value)} /></div>
          </div>

          <div className="bg-gray-300 font-bold p-2 border-b-2 border-black text-center uppercase tracking-wider text-blue-900">Plano de Ação (Para eliminar a Causa Raiz)</div>
          <div className="flex border-b-2 border-black font-bold text-center bg-gray-100 text-xs uppercase">
            <div className="w-1/2 border-r-2 border-black p-2 flex items-center justify-center">O Quê (Ação a ser executada)</div>
            <div className="w-1/6 border-r-2 border-black p-2 flex items-center justify-center">Quem (Responsável)</div>
            <div className="w-1/6 border-r-2 border-black p-2 flex items-center justify-center">Quando (Previsto)</div>
            <div className="w-1/6 p-2 flex items-center justify-center">Status / Realizado</div>
          </div>
          <div className="flex border-b-2 border-black bg-white">
            <div className="w-1/2 border-r-2 border-black"><textarea className="w-full h-full p-3 focus:outline-none resize-none" value={rnc.actionPlan} onChange={(e) => handleRncChange(rnc.id, 'actionPlan', e.target.value)}></textarea></div>
            <div className="w-1/6 border-r-2 border-black"><input type="text" className="w-full h-full p-3 focus:outline-none text-center" value={rnc.responsible} onChange={(e) => handleRncChange(rnc.id, 'responsible', e.target.value)} /></div>
            <div className="w-1/6 border-r-2 border-black"><input type="date" className="w-full h-full p-3 focus:outline-none text-center text-xs" value={rnc.actionPlanDatePrev} onChange={(e) => handleRncChange(rnc.id, 'actionPlanDatePrev', e.target.value)} /></div>
            <div className="w-1/6"><input type="date" className="w-full h-full p-3 focus:outline-none text-center text-xs" value={rnc.actionPlanDateReal} onChange={(e) => handleRncChange(rnc.id, 'actionPlanDateReal', e.target.value)} /></div>
          </div>

          <div className="bg-gray-300 font-bold p-2 border-b-2 border-black text-center uppercase tracking-wider">Análise Crítica da Eficácia das Ações (Validação SGQ)</div>
          <div className="p-4 border-b-2 border-black bg-white">
            <div className="font-bold text-gray-700 mb-2 uppercase text-xs tracking-wider">Evidências Coletadas após Implementação:</div>
            <textarea className="w-full border-2 border-gray-200 p-3 focus:outline-none bg-gray-50 rounded" rows="3" value={rnc.evidence} onChange={(e) => handleRncChange(rnc.id, 'evidence', e.target.value)}></textarea>
          </div>

          <div className="flex border-b-2 border-black bg-gray-50">
            <div className="w-1/2 border-r-2 border-black p-4 flex items-center justify-center gap-6">
              <span className="font-bold uppercase text-sm">Ação foi Eficaz?</span>
              <label className="flex items-center gap-2 cursor-pointer font-bold"><input type="radio" className="w-4 h-4" name={`eficaz-${rnc.id}`} checked={rnc.effective === 'Sim'} onChange={() => handleRncChange(rnc.id, 'effective', 'Sim')} /> SIM</label>
              <label className="flex items-center gap-2 cursor-pointer font-bold"><input type="radio" className="w-4 h-4" name={`eficaz-${rnc.id}`} checked={rnc.effective === 'Não'} onChange={() => handleRncChange(rnc.id, 'effective', 'Não')} /> NÃO</label>
            </div>
            <div className="w-1/2 p-4 flex items-center justify-center gap-6">
              <span className="font-bold uppercase text-sm">Necessário Nova RNC?</span>
              <label className="flex items-center gap-2 cursor-pointer font-bold"><input type="radio" className="w-4 h-4" name={`novo-${rnc.id}`} checked={rnc.newRnc === 'Sim'} onChange={() => handleRncChange(rnc.id, 'newRnc', 'Sim')} /> SIM</label>
              <label className="flex items-center gap-2 cursor-pointer font-bold"><input type="radio" className="w-4 h-4" name={`novo-${rnc.id}`} checked={rnc.newRnc === 'Não'} onChange={() => handleRncChange(rnc.id, 'newRnc', 'Não')} /> NÃO</label>
            </div>
          </div>

          <div className="flex bg-white">
            <div className="w-1/2 border-r-2 border-black p-4 flex items-center gap-4">
              <span className="font-bold uppercase text-sm whitespace-nowrap">Data Fechamento:</span>
              <input type="date" className="border-b-2 border-gray-400 focus:outline-none flex-1 bg-transparent px-2 py-1 font-bold" value={rnc.closeDate} onChange={(e) => handleRncChange(rnc.id, 'closeDate', e.target.value)} />
            </div>
            <div className="w-1/2 p-4 flex items-center gap-4">
              <span className="font-bold uppercase text-sm whitespace-nowrap">Validador SGQ:</span>
              <input type="text" className="border-b-2 border-gray-400 focus:outline-none flex-1 bg-transparent px-2 py-1 font-bold" value={rnc.closeResp} onChange={(e) => handleRncChange(rnc.id, 'closeResp', e.target.value)} />
            </div>
          </div>

          <div className="bg-slate-900 border-t-2 border-black p-4 flex justify-between items-center gap-3 print:hidden">
            <button
              onClick={() => requestConfirm({
                title: 'Excluir RNC',
                message: `Tem certeza que deseja excluir permanentemente a RNC Nº ${rnc.id}? Essa ação não pode ser desfeita.`,
                confirmLabel: 'Sim, excluir',
                danger: true,
                onConfirm: async () => {
                    setRncs((prev) => prev.filter((r) => r.id !== rnc.id));
                    if (rnc.dbId) {
                      try { await supaFetch(`sga_rncs?id=eq.${rnc.dbId}`, { method: 'DELETE' }); } catch {}
                    }
                  }
              })}
              className="text-rose-400 hover:text-rose-300 font-bold text-sm flex items-center gap-1.5 transition"
            >
              <Trash2 size={15} /> Excluir RNC
            </button>
            <button
              onClick={() => {
                if (rnc.status === 'Aberta' && !rnc.evidence.trim()) {
                  notify('Faltam evidências', 'Registre as evidências coletadas após a implementação antes de fechar esta RNC.');
                  return;
                }
                requestConfirm({
                  title: rnc.status === 'Aberta' ? 'Confirmar fechamento da RNC' : 'Confirmar reabertura da RNC',
                  message: rnc.status === 'Aberta'
                    ? `Tem certeza que deseja registrar o fechamento da RNC Nº ${rnc.id}? Essa ação marca a não conformidade como tratada e validada pelo SGQ.`
                    : `Tem certeza que deseja reabrir a RNC Nº ${rnc.id}? Ela voltará para a lista de pendências em aberto.`,
                  confirmLabel: rnc.status === 'Aberta' ? 'Sim, fechar RNC' : 'Sim, reabrir RNC',
                  danger: rnc.status !== 'Aberta',
                  onConfirm: () => handleRncChange(rnc.id, 'status', rnc.status === 'Aberta' ? 'Fechada' : 'Aberta')
                });
              }}
              className="bg-indigo-600 text-white font-bold px-8 py-3 rounded-xl flex items-center gap-2 hover:bg-indigo-500 transition shadow-lg"
            >
              <Save size={20} /> {rnc.status === 'Aberta' ? 'Registrar Fechamento' : 'Reabrir RNC'}
            </button>
          </div>
        </div>
      ))}
    </div>
  );
  };

  const renderGestao = () => (
    <div className="space-y-6 animate-fade-in">
      <div className="flex justify-between items-center mb-6">
        <div>
          <h2 className="text-2xl font-black text-slate-900">Repositório de Auditorias</h2>
          <p className="text-slate-500 mt-1">Histórico consolidado das unidades de Minas Gerais</p>
        </div>

        <div className="flex items-center gap-3">
          <div className="flex bg-white border border-slate-200 rounded-xl px-4 py-2.5 shadow-sm w-80 focus-within:ring-2 focus-within:ring-indigo-500 focus-within:border-indigo-500">
            <Search className="text-slate-400 mr-2" size={20} />
            <input
              type="text"
              placeholder="Buscar por Nº RAI, Setor ou Auditor..."
              className="w-full outline-none text-sm bg-transparent"
              value={historySearch}
              onChange={(e) => setHistorySearch(e.target.value)}
            />
          </div>
          {savedAudits.length > 0 && (
            <button onClick={exportHistoryExcel} className="flex items-center gap-2 bg-white border border-slate-200 text-emerald-700 px-4 py-2.5 rounded-xl font-bold hover:bg-emerald-50 transition text-sm shadow-sm whitespace-nowrap">
              <FileSpreadsheet size={16} /> Exportar Excel
            </button>
          )}
        </div>
      </div>

      <div className="flex gap-2">
        {[
          { key: 'Todas', label: 'Todas' },
          { key: 'ComNC', label: 'Com NC' },
          { key: 'SemNC', label: 'Sem NC' }
        ].map((opt) => (
          <button
            key={opt.key}
            onClick={() => setHistoryNcFilter(opt.key)}
            className={`px-4 py-1.5 rounded-full text-xs font-bold border transition ${
              historyNcFilter === opt.key
                ? 'bg-indigo-600 text-white border-indigo-600 shadow-sm'
                : 'bg-white text-slate-500 border-slate-200 hover:bg-slate-50'
            }`}
          >
            {opt.label}
          </button>
        ))}
      </div>

      {savedAudits.length === 0 ? (
        <div className="bg-white border border-slate-200 rounded-2xl p-16 text-center text-slate-500 shadow-sm font-medium flex flex-col items-center justify-center">
          <FolderTree size={48} className="text-slate-300 mb-4" />
          Nenhum registro encontrado. Realize e salve uma auditoria para popular o painel.
        </div>
      ) : filteredAudits.length === 0 ? (
        <div className="bg-white border border-slate-200 rounded-2xl p-16 text-center text-slate-500 shadow-sm font-medium flex flex-col items-center justify-center">
          <Search size={48} className="text-slate-300 mb-4" />
          Nenhuma auditoria encontrada para "{historySearch}".
        </div>
      ) : (
        <div className="bg-white shadow-sm border border-slate-200 rounded-2xl overflow-hidden">
          <table className="w-full text-left border-collapse font-sans text-sm">
            <thead>
              <tr className="bg-slate-50 text-slate-500 font-bold border-b border-slate-200 uppercase tracking-wider text-xs">
                <th className="p-4 text-center w-24">Nº RAI</th>
                <th className="p-4 text-center w-32">DATA</th>
                <th className="p-4">UNIDADE</th>
                <th className="p-4">ÁREA / PROCESSO</th>
                <th className="p-4">AUDITOR LÍDER</th>
                <th className="p-4 text-center w-24">NCs</th>
                <th className="p-4 text-center w-32">AÇÕES</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 bg-white">
              {pagedAudits.map((audit) => (
                <tr key={audit.id} className="hover:bg-indigo-50/40 transition-colors">
                  <td className="p-4 text-center font-black text-slate-800 font-mono">{audit.raiNumber}</td>
                  <td className="p-4 text-center text-slate-500">{audit.date}</td>
                  <td className="p-4 font-bold text-indigo-700">{audit.branch}</td>
                  <td className="p-4 font-bold text-slate-700 uppercase">{audit.sector}</td>
                  <td className="p-4 text-slate-500">{audit.auditor}</td>
                  <td className="p-4 text-center">
                    <span className={`px-3 py-1 rounded-full text-xs font-black shadow-sm border ${audit.ncCount > 0 ? 'bg-rose-50 text-rose-700 border-rose-200' : 'bg-emerald-50 text-emerald-700 border-emerald-200'}`}>
                      {audit.ncCount}
                    </span>
                  </td>
                  <td className="p-4 text-center">
                    <div className="flex items-center justify-center gap-2">
                      <button
                        onClick={() => loadAudit(audit)}
                        disabled={loadingAuditId === audit.id}
                        className="bg-white text-indigo-600 border border-indigo-200 hover:bg-indigo-600 hover:text-white px-4 py-1.5 rounded-lg text-xs font-bold transition-all shadow-sm disabled:opacity-50 disabled:cursor-wait"
                      >
                        {loadingAuditId === audit.id ? '...' : 'VER'}
                      </button>
                      {isAdmin && (
                        <button
                          onClick={() => requestConfirm({
                            title: 'Editar auditoria',
                            message: `Deseja abrir a auditoria RAI Nº ${audit.raiNumber} para edição? O checklist e o relatório poderão ser alterados e salvos novamente.`,
                            confirmLabel: 'Sim, editar',
                            danger: false,
                            onConfirm: () => loadAudit(audit, true)
                          })}
                          className="bg-white text-amber-600 border border-amber-200 hover:bg-amber-500 hover:text-white p-1.5 rounded-lg transition-all shadow-sm"
                          title="Editar auditoria"
                        >
                          <Edit3 size={14} />
                        </button>
                      )}
                      {isAdmin && (
                        <button
                          onClick={() => requestConfirm({
                            title: 'Excluir auditoria',
                            message: `Tem certeza que deseja excluir permanentemente a auditoria RAI Nº ${audit.raiNumber} (${audit.sector})? Essa ação não pode ser desfeita.`,
                            confirmLabel: 'Sim, excluir',
                            danger: true,
                            onConfirm: () => deleteAudit(audit.id)
                          })}
                          aria-label="Excluir auditoria"
                          className="bg-white text-rose-600 border border-rose-200 hover:bg-rose-600 hover:text-white p-1.5 rounded-lg transition-all shadow-sm"
                        >
                          <Trash2 size={14} />
                        </button>
                      )}
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          <div className="bg-slate-50 border-t border-slate-200 p-4 text-xs text-slate-500 flex justify-between items-center">
            <span>
              Exibindo {pagedAudits.length} de {filteredAudits.length} registro(s)
              {historySearch ? ` (filtrado de ${savedAudits.length} no total)` : ''}.
            </span>
            <div className="flex items-center gap-3">
              <span className="text-slate-400">Página {historyPage} de {historyTotalPages}</span>
              <div className="flex gap-2">
                <button
                  onClick={() => setHistoryPage((p) => Math.max(1, p - 1))}
                  disabled={historyPage <= 1}
                  className="px-3 py-1 border border-slate-300 rounded-lg bg-white hover:bg-slate-100 disabled:opacity-50 disabled:cursor-not-allowed"
                >
                  Anterior
                </button>
                <button
                  onClick={() => setHistoryPage((p) => Math.min(historyTotalPages, p + 1))}
                  disabled={historyPage >= historyTotalPages}
                  className="px-3 py-1 border border-slate-300 rounded-lg bg-white hover:bg-slate-100 disabled:opacity-50 disabled:cursor-not-allowed"
                >
                  Próximo
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );

  // ---- Calendário de programação de auditorias ----

  const renderEventChip = (ev, compact = false) => (
    <div
      key={ev.id}
      draggable
      onDragStart={() => setDraggedEventId(ev.id)}
      onDragEnd={() => { setDraggedEventId(null); setDragOverDateKey(null); }}
      onClick={(e) => { e.stopPropagation(); setEventDraft(ev); }}
      className={`group flex items-center gap-1 bg-indigo-600 hover:bg-indigo-700 text-white rounded-md px-1.5 py-1 cursor-grab active:cursor-grabbing transition ${compact ? 'text-[10px]' : 'text-xs'}`}
      title={`${ev.sector} · ${formatTimeRange(ev.time, ev.durationMin)} · ${ev.auditor}`}
    >
      <GripVertical size={10} className="shrink-0 opacity-50 group-hover:opacity-90" />
      <span className="font-bold truncate">{ev.time}</span>
      <span className="truncate">{ev.sector}</span>
    </div>
  );

  const renderMonthView = () => {
    const year = calendarAnchor.getFullYear();
    const month = calendarAnchor.getMonth();
    const grid = getMonthGrid(year, month);
    const todayKey = toDateKey(new Date());

    return (
      <div className="bg-white rounded-2xl shadow-sm border border-slate-100 overflow-hidden">
        <div className="grid grid-cols-7 bg-slate-50 border-b border-slate-200">
          {WEEKDAY_LABELS.map((d) => (
            <div key={d} className="p-3 text-center text-[11px] font-bold text-slate-400 uppercase tracking-wider">{d}</div>
          ))}
        </div>
        <div className="grid grid-cols-7">
          {grid.map((day) => {
            const dateKey = toDateKey(day);
            const inMonth = day.getMonth() === month;
            const isToday = dateKey === todayKey;
            const dayEvents = auditEvents.filter((e) => e.dateKey === dateKey).sort((a, b) => a.time.localeCompare(b.time));
            const isDragOver = dragOverDateKey === dateKey;

            return (
              <div
                key={dateKey}
                onDragOver={(e) => { e.preventDefault(); setDragOverDateKey(dateKey); }}
                onDragLeave={() => setDragOverDateKey((prev) => (prev === dateKey ? null : prev))}
                onDrop={(e) => { e.preventDefault(); handleDropOnDate(dateKey); }}
                onClick={() => setEventDraft({ id: null, dateKey, time: '09:00', durationMin: 60, sector: selectedSector, auditor: AUDITORES[0], branch: selectedBranch, notes: '' })}
                className={`min-h-[100px] border-b border-r border-slate-100 p-1.5 cursor-pointer transition-colors ${
                  inMonth ? 'bg-white hover:bg-slate-50' : 'bg-slate-50/50 text-slate-300'
                } ${isDragOver ? 'bg-indigo-50 ring-2 ring-inset ring-indigo-400' : ''}`}
              >
                <div className={`text-xs font-bold mb-1 w-6 h-6 flex items-center justify-center rounded-full ${isToday ? 'bg-indigo-600 text-white' : inMonth ? 'text-slate-600' : 'text-slate-300'}`}>
                  {day.getDate()}
                </div>
                <div className="space-y-1">
                  {dayEvents.map((ev) => renderEventChip(ev, true))}
                </div>
              </div>
            );
          })}
        </div>
      </div>
    );
  };

  const renderWeekView = () => {
    const days = getWeekGrid(calendarAnchor);
    const todayKey = toDateKey(new Date());

    return (
      <div className="bg-white rounded-2xl shadow-sm border border-slate-100 overflow-hidden">
        <div className="grid grid-cols-7 bg-slate-50 border-b border-slate-200">
          {days.map((day) => {
            const dateKey = toDateKey(day);
            const isToday = dateKey === todayKey;
            return (
              <div key={dateKey} className="p-3 text-center border-r border-slate-100 last:border-r-0">
                <div className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">{WEEKDAY_LABELS[day.getDay()]}</div>
                <div className={`mt-1 text-sm font-black w-7 h-7 mx-auto flex items-center justify-center rounded-full ${isToday ? 'bg-indigo-600 text-white' : 'text-slate-700'}`}>
                  {day.getDate()}
                </div>
              </div>
            );
          })}
        </div>
        <div className="grid grid-cols-7 min-h-[420px]">
          {days.map((day) => {
            const dateKey = toDateKey(day);
            const dayEvents = auditEvents.filter((e) => e.dateKey === dateKey).sort((a, b) => a.time.localeCompare(b.time));
            const isDragOver = dragOverDateKey === dateKey;
            return (
              <div
                key={dateKey}
                onDragOver={(e) => { e.preventDefault(); setDragOverDateKey(dateKey); }}
                onDragLeave={() => setDragOverDateKey((prev) => (prev === dateKey ? null : prev))}
                onDrop={(e) => { e.preventDefault(); handleDropOnDate(dateKey); }}
                onClick={() => setEventDraft({ id: null, dateKey, time: '09:00', durationMin: 60, sector: selectedSector, auditor: AUDITORES[0], branch: selectedBranch, notes: '' })}
                className={`border-r border-slate-100 last:border-r-0 p-2 space-y-1.5 cursor-pointer hover:bg-slate-50 transition-colors ${isDragOver ? 'bg-indigo-50 ring-2 ring-inset ring-indigo-400' : ''}`}
              >
                {dayEvents.map((ev) => renderEventChip(ev, false))}
              </div>
            );
          })}
        </div>
      </div>
    );
  };

  const renderCalendario = () => (
    <div className="space-y-6 animate-fade-in">
      <div className="flex flex-col sm:flex-row sm:justify-between sm:items-center gap-4 mb-2">
        <div>
          <h2 className="text-2xl font-black text-slate-900">Calendário de Auditorias</h2>
          <p className="text-slate-500 text-sm mt-1">Programe as auditorias por setor — arraste um item para reagendar</p>
        </div>
        <button
          onClick={() => setEventDraft({ id: null, dateKey: toDateKey(new Date()), time: '09:00', durationMin: 60, sector: selectedSector, auditor: AUDITORES[0], branch: selectedBranch, notes: '' })}
          className="bg-indigo-600 hover:bg-indigo-700 text-white px-5 py-3 rounded-xl font-bold transition shadow-lg shadow-indigo-600/25 flex items-center gap-2 self-start"
        >
          <Plus size={18} strokeWidth={2.5} /> Programar Auditoria
        </button>
      </div>

      <div className="flex flex-wrap justify-between items-center gap-3 bg-white rounded-2xl shadow-sm border border-slate-100 p-3">
        <div className="flex items-center gap-2">
          <button
            onClick={() => setCalendarAnchor((d) => (calendarView === 'month' ? new Date(d.getFullYear(), d.getMonth() - 1, 1) : addDays(d, -7)))}
            className="p-2 rounded-lg hover:bg-slate-100 text-slate-600 transition"
          >
            <ChevronLeft size={18} />
          </button>
          <span className="font-bold text-slate-800 min-w-[160px] text-center">
            {calendarView === 'month'
              ? `${MONTH_LABELS[calendarAnchor.getMonth()]} de ${calendarAnchor.getFullYear()}`
              : `Semana de ${startOfWeek(calendarAnchor).toLocaleDateString('pt-BR', { day: '2-digit', month: '2-digit' })}`}
          </span>
          <button
            onClick={() => setCalendarAnchor((d) => (calendarView === 'month' ? new Date(d.getFullYear(), d.getMonth() + 1, 1) : addDays(d, 7)))}
            className="p-2 rounded-lg hover:bg-slate-100 text-slate-600 transition"
          >
            <ChevronRight size={18} />
          </button>
          <button
            onClick={() => setCalendarAnchor(new Date())}
            className="ml-2 text-xs font-bold text-indigo-600 hover:underline"
          >
            Hoje
          </button>
        </div>
        <div className="flex bg-slate-100 rounded-xl p-1">
          <button
            onClick={() => setCalendarView('month')}
            className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold transition ${calendarView === 'month' ? 'bg-white shadow-sm text-indigo-700' : 'text-slate-500'}`}
          >
            <CalendarDays size={14} /> Mês
          </button>
          <button
            onClick={() => setCalendarView('week')}
            className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold transition ${calendarView === 'week' ? 'bg-white shadow-sm text-indigo-700' : 'text-slate-500'}`}
          >
            <CalendarRange size={14} /> Semana
          </button>
        </div>
      </div>

      {calendarView === 'month' ? renderMonthView() : renderWeekView()}

      <p className="text-xs text-slate-400 text-center print:hidden">
        Clique em um dia vazio para programar uma nova auditoria, ou em um item existente para editar. Arraste para reagendar.
      </p>
    </div>
  );


  const renderDocumentos = () => (
    <div className="space-y-6 animate-fade-in">
      <div className="flex flex-col sm:flex-row sm:justify-between sm:items-center gap-4 mb-2">
        <div>
          <h2 className="text-2xl font-black text-slate-900">Documentos por Setor</h2>
          <p className="text-slate-500 text-sm mt-1">
            {documentos.length} documento(s) cadastrado(s)
            {docsVencidos > 0 && <span className="text-rose-600 font-bold ml-2">· {docsVencidos} vencido(s)</span>}
            {docsProximos > 0 && <span className="text-amber-600 font-bold ml-2">· {docsProximos} vence em 30 dias</span>}
          </p>
        </div>
        <button
          onClick={() => setDocDraft({ setor: docSectorFilter !== 'TODOS' ? docSectorFilter : SECTORS[0], nome: '', tipo: DOC_TIPOS[0], url: '', responsavel: '', validade: '', descricao: '' })}
          className="bg-indigo-600 hover:bg-indigo-700 text-white px-5 py-3 rounded-xl font-bold transition shadow-lg shadow-indigo-600/25 flex items-center gap-2 self-start"
        >
          <Plus size={18} strokeWidth={2.5} /> Adicionar Documento
        </button>
      </div>

      {/* Alertas de vencimento */}
      {(docsVencidos > 0 || docsProximos > 0) && (
        <div className="space-y-2">
          {docsVencidos > 0 && (
            <div className="bg-rose-50 border border-rose-200 text-rose-800 rounded-xl p-4 flex items-center gap-3 text-sm">
              <FileWarning size={18} className="shrink-0" />
              <span><span className="font-bold">{docsVencidos} documento(s) vencido(s)</span> — atualize ou renove urgentemente.</span>
            </div>
          )}
          {docsProximos > 0 && (
            <div className="bg-amber-50 border border-amber-200 text-amber-800 rounded-xl p-4 flex items-center gap-3 text-sm">
              <AlertCircle size={18} className="shrink-0" />
              <span><span className="font-bold">{docsProximos} documento(s)</span> com vencimento nos próximos 30 dias.</span>
            </div>
          )}
        </div>
      )}

      {/* Filtros — admin vê todos os setores, usuário de setor vê só o seu */}
      <div className="flex flex-wrap gap-3 items-center bg-white rounded-2xl shadow-sm border border-slate-100 p-4">
        {isAdmin ? (
          <>
            <div className="flex items-center gap-2 text-slate-500">
              <Filter size={15} />
              <span className="text-xs font-bold uppercase tracking-wider">Filtrar por setor:</span>
            </div>
            <div className="flex flex-wrap gap-1.5">
              {['TODOS', ...SECTORS].map((s) => (
                <button
                  key={s}
                  onClick={() => setDocSectorFilter(s)}
                  className={`px-3 py-1 rounded-full text-xs font-bold border transition ${
                    docSectorFilter === s
                      ? 'bg-indigo-600 text-white border-indigo-600'
                      : 'bg-white text-slate-500 border-slate-200 hover:bg-slate-50'
                  }`}
                >
                  {s === 'TODOS' ? 'Todos' : s}
                  {s !== 'TODOS' && (
                    <span className="ml-1 opacity-60">({documentos.filter((d) => d.setor === s).length})</span>
                  )}
                </button>
              ))}
            </div>
          </>
        ) : (
          <div className="flex items-center gap-2">
            <Folder size={15} className="text-indigo-600" />
            <span className="text-sm font-bold text-slate-700">
              Setor: <span className="text-indigo-600">{userSetores[0]}</span>
            </span>
            <span className="text-xs text-slate-400 ml-2">({documentos.filter((d) => userSetores.includes(d.setor)).length} documento(s))</span>
          </div>
        )}
        <div className="ml-auto flex items-center gap-2 bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 min-w-[220px] focus-within:ring-2 focus-within:ring-indigo-500">
          <Search size={14} className="text-slate-400 shrink-0" />
          <input
            type="text"
            placeholder="Buscar documento, tipo..."
            value={docSearch}
            onChange={(e) => setDocSearch(e.target.value)}
            className="bg-transparent text-sm outline-none w-full"
          />
        </div>
      </div>

      {/* Lista de documentos */}
      {filteredDocumentos.length === 0 ? (
        <div className="bg-white border border-slate-200 rounded-2xl p-16 text-center text-slate-400 shadow-sm flex flex-col items-center justify-center gap-4">
          <Folder size={56} className="text-slate-200" />
          <div>
            <p className="font-bold text-slate-500">Nenhum documento encontrado</p>
            <p className="text-sm mt-1">
              {docSearch || docSectorFilter !== 'TODOS'
                ? 'Tente ajustar os filtros de busca.'
                : 'Clique em "Adicionar Documento" para registrar o primeiro documento.'}
            </p>
          </div>
        </div>
      ) : (
        <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-4">
          {filteredDocumentos.map((doc) => {
            const status = getDocStatus(doc.validade);
            const statusColors = {
              vencido: { bg: 'bg-rose-50', border: 'border-rose-200', badge: 'bg-rose-100 text-rose-700', dot: 'bg-rose-500', label: 'Vencido' },
              proximo: { bg: 'bg-amber-50', border: 'border-amber-200', badge: 'bg-amber-100 text-amber-700', dot: 'bg-amber-400', label: 'Vence em breve' },
              ok:      { bg: 'bg-white', border: 'border-slate-100', badge: 'bg-emerald-50 text-emerald-700', dot: 'bg-emerald-500', label: 'Válido' },
            };
            const sc = statusColors[status] || statusColors.ok;
            return (
              <div key={doc.id} className={`${sc.bg} border ${sc.border} rounded-2xl p-5 shadow-sm hover:shadow-md transition-shadow flex flex-col gap-3 group`}>
                <div className="flex items-start justify-between gap-2">
                  <div className="flex-1 min-w-0">
                    <p className="font-black text-slate-800 text-sm leading-snug truncate">{doc.nome}</p>
                    <p className="text-xs text-slate-500 mt-0.5">{doc.setor}</p>
                  </div>
                  <div className="flex gap-1 shrink-0 opacity-0 group-hover:opacity-100 transition">
                    <button onClick={() => { setDocDraft({ ...doc }); setDocUploadFile(null); setDocUploadError(''); }} className="p-1.5 rounded-lg hover:bg-indigo-50 text-indigo-600 transition" title="Editar">
                      <Edit3 size={14} />
                    </button>
                    <button
                      onClick={() => requestConfirm({
                        title: 'Excluir documento',
                        message: `Tem certeza que deseja excluir "${doc.nome}"?`,
                        confirmLabel: 'Sim, excluir',
                        danger: true,
                        onConfirm: () => deleteDocumento(doc)
                      })}
                      className="p-1.5 rounded-lg hover:bg-rose-50 text-rose-500 transition" title="Excluir"
                    >
                      <Trash2 size={14} />
                    </button>
                  </div>
                </div>

                <div className="flex flex-wrap gap-1.5">
                  <span className="bg-indigo-50 text-indigo-700 text-[10px] font-bold px-2 py-0.5 rounded-full">{doc.tipo}</span>
                  {status && (
                    <span className={`${sc.badge} text-[10px] font-bold px-2 py-0.5 rounded-full flex items-center gap-1`}>
                      <span className={`w-1.5 h-1.5 rounded-full ${sc.dot}`}></span>
                      {sc.label}
                    </span>
                  )}
                </div>

                {doc.descricao && (
                  <p className="text-xs text-slate-500 leading-relaxed line-clamp-2">{doc.descricao}</p>
                )}

                <div className="mt-auto space-y-1 text-xs text-slate-500">
                  {doc.storage_name && (
                    <div className="flex items-center gap-1.5 text-indigo-600 font-medium">
                      <FileBadge size={11} />
                      <span className="truncate">{doc.storage_name}</span>
                      {doc.storage_size && <span className="shrink-0 text-slate-400">· {formatFileSize(doc.storage_size)}</span>}
                    </div>
                  )}
                  {doc.responsavel && (
                    <div className="flex items-center gap-1.5"><User size={11} /> <span>{doc.responsavel}</span></div>
                  )}
                  {doc.validade && (
                    <div className="flex items-center gap-1.5">
                      <Clock size={11} />
                      <span>Validade: {new Date(doc.validade + 'T00:00:00').toLocaleDateString('pt-BR')}</span>
                    </div>
                  )}
                </div>

                {doc.url && (
                  <a
                    href={doc.url.startsWith('http') ? doc.url : `file://${doc.url}`}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="flex items-center gap-2 text-indigo-600 hover:text-indigo-800 text-xs font-bold border border-indigo-100 bg-indigo-50 hover:bg-indigo-100 rounded-lg px-3 py-2 transition mt-1"
                  >
                    <ExternalLink size={13} />
                    {doc.storage_name ? 'Baixar arquivo' : 'Abrir documento'}
                  </a>
                )}
              </div>
            );
          })}
        </div>
      )}

      {/* Modal de criação/edição de documento */}
      {docDraft && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4" onClick={() => { setDocDraft(null); setDocUploadFile(null); setDocUploadError(''); }}>
          <div className="bg-white rounded-2xl shadow-2xl max-w-lg w-full p-6 max-h-[90vh] overflow-y-auto" onClick={(e) => e.stopPropagation()}>
            <div className="flex items-start justify-between mb-5">
              <div>
                <h3 className="text-lg font-black text-slate-900">{docDraft.id ? 'Editar Documento' : 'Adicionar Documento'}</h3>
                <p className="text-xs text-slate-500 mt-1">Registre o documento e o link/caminho de acesso na rede</p>
              </div>
              <button onClick={() => { setDocDraft(null); setDocUploadFile(null); setDocUploadError(''); }} className="text-slate-400 hover:text-slate-600"><X size={18} /></button>
            </div>

            <div className="space-y-4">
              <div>
                <label className="text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5 block">Setor *</label>
                <select
                  value={docDraft.setor}
                  onChange={(e) => setDocDraft({ ...docDraft, setor: e.target.value })}
                  className="w-full border border-slate-200 rounded-xl p-2.5 text-sm focus:ring-2 focus:ring-indigo-500"
                >
                  {SECTORS.map((s) => <option key={s} value={s}>{s}</option>)}
                </select>
              </div>

              <div>
                <label className="text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5 block">Nome do Documento *</label>
                <input
                  type="text"
                  placeholder="Ex.: PROC027 - Controle de Documentos"
                  value={docDraft.nome}
                  onChange={(e) => setDocDraft({ ...docDraft, nome: e.target.value })}
                  className="w-full border border-slate-200 rounded-xl p-2.5 text-sm focus:ring-2 focus:ring-indigo-500"
                />
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5 block">Tipo *</label>
                  <select
                    value={docDraft.tipo}
                    onChange={(e) => setDocDraft({ ...docDraft, tipo: e.target.value })}
                    className="w-full border border-slate-200 rounded-xl p-2.5 text-sm focus:ring-2 focus:ring-indigo-500"
                  >
                    {DOC_TIPOS.map((t) => <option key={t} value={t}>{t}</option>)}
                  </select>
                </div>
                <div>
                  <label className="text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5 block">Validade</label>
                  <input
                    type="date"
                    value={docDraft.validade}
                    onChange={(e) => setDocDraft({ ...docDraft, validade: e.target.value })}
                    className="w-full border border-slate-200 rounded-xl p-2.5 text-sm focus:ring-2 focus:ring-indigo-500"
                  />
                </div>
              </div>

              <div>
                {/* Upload de arquivo OU link externo — os dois são opcionais mas pelo menos um é útil */}
                <label className="text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5 block">Arquivo do Documento</label>

                {/* Arquivo já carregado (edição) */}
                {docDraft.storage_name && !docUploadFile && (
                  <div className="flex items-center gap-3 bg-indigo-50 border border-indigo-200 rounded-xl p-3 mb-2">
                    <FileBadge size={18} className="text-indigo-600 shrink-0" />
                    <div className="flex-1 min-w-0">
                      <p className="text-sm font-bold text-slate-700 truncate">{docDraft.storage_name}</p>
                      <p className="text-xs text-slate-500">{formatFileSize(docDraft.storage_size)}</p>
                    </div>
                    <button
                      onClick={() => setDocDraft({ ...docDraft, storage_path: null, storage_name: null, storage_size: null, storage_mime: null })}
                      className="text-slate-400 hover:text-rose-500 transition shrink-0"
                      title="Remover arquivo"
                    >
                      <X size={16} />
                    </button>
                  </div>
                )}

                {/* Arquivo selecionado para upload (antes de salvar) */}
                {docUploadFile && (
                  <div className="flex items-center gap-3 bg-emerald-50 border border-emerald-200 rounded-xl p-3 mb-2">
                    <Upload size={18} className="text-emerald-600 shrink-0" />
                    <div className="flex-1 min-w-0">
                      <p className="text-sm font-bold text-slate-700 truncate">{docUploadFile.name}</p>
                      <p className="text-xs text-slate-500">{formatFileSize(docUploadFile.size)} · Pronto para enviar</p>
                    </div>
                    <button onClick={() => setDocUploadFile(null)} className="text-slate-400 hover:text-rose-500 transition shrink-0">
                      <X size={16} />
                    </button>
                  </div>
                )}

                {/* Botão de selecionar arquivo */}
                {!docUploadFile && (
                  <label className="flex items-center gap-3 border-2 border-dashed border-slate-200 hover:border-indigo-400 rounded-xl p-4 cursor-pointer transition group">
                    <Upload size={20} className="text-slate-400 group-hover:text-indigo-500 transition shrink-0" />
                    <div>
                      <p className="text-sm font-bold text-slate-600 group-hover:text-indigo-600 transition">
                        {docDraft.storage_name ? 'Trocar arquivo' : 'Carregar arquivo'}
                      </p>
                      <p className="text-xs text-slate-400">PDF, Word, Excel, imagem — até 50 MB</p>
                    </div>
                    <input
                      type="file"
                      className="hidden"
                      accept=".pdf,.doc,.docx,.xls,.xlsx,.png,.jpg,.jpeg,.txt,.csv,.zip"
                      onChange={(e) => {
                        const f = e.target.files?.[0];
                        if (f) { setDocUploadFile(f); setDocUploadError(''); }
                      }}
                    />
                  </label>
                )}

                {/* Separador OU */}
                <div className="flex items-center gap-3 my-3">
                  <div className="flex-1 h-px bg-slate-200"></div>
                  <span className="text-xs text-slate-400 font-bold">OU cole um link externo</span>
                  <div className="flex-1 h-px bg-slate-200"></div>
                </div>

                <input
                  type="text"
                  placeholder="https://sharepoint.com/... ou \\servidor\pasta\arquivo.pdf"
                  value={docDraft.url || ''}
                  onChange={(e) => setDocDraft({ ...docDraft, url: e.target.value })}
                  className="w-full border border-slate-200 rounded-xl p-2.5 text-sm focus:ring-2 focus:ring-indigo-500 font-mono"
                />
                <p className="text-[10px] text-slate-400 mt-1">SharePoint, OneDrive, Google Drive, caminho de rede</p>
              </div>

              <div>
                <label className="text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5 block">Responsável</label>
                <input
                  type="text"
                  placeholder="Nome do responsável pelo documento"
                  value={docDraft.responsavel}
                  onChange={(e) => setDocDraft({ ...docDraft, responsavel: e.target.value })}
                  className="w-full border border-slate-200 rounded-xl p-2.5 text-sm focus:ring-2 focus:ring-indigo-500"
                />
              </div>

              <div>
                <label className="text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5 block">Descrição / Observações</label>
                <textarea
                  rows="2"
                  placeholder="Breve descrição do propósito do documento..."
                  value={docDraft.descricao}
                  onChange={(e) => setDocDraft({ ...docDraft, descricao: e.target.value })}
                  className="w-full border border-slate-200 rounded-xl p-2.5 text-sm focus:ring-2 focus:ring-indigo-500 resize-none"
                />
              </div>

              {docUploadError && (
                <div className="bg-rose-50 border border-rose-200 text-rose-700 rounded-xl p-3 text-sm flex items-start gap-2">
                  <AlertCircle size={16} className="mt-0.5 shrink-0" />
                  {docUploadError}
                </div>
              )}
            </div>

            <div className="flex justify-end gap-3 mt-6">
              <button
                onClick={() => { setDocDraft(null); setDocUploadFile(null); setDocUploadError(''); }}
                className="px-4 py-2.5 rounded-xl font-bold text-sm text-slate-600 hover:bg-slate-100 transition"
                disabled={docUploading}
              >
                Cancelar
              </button>
              <button
                onClick={() => {
                  if (!docDraft.nome.trim() || !docDraft.setor) {
                    setDocUploadError('Preencha pelo menos o nome e o setor do documento.');
                    return;
                  }
                  if (!docUploadFile && !docDraft.url && !docDraft.storage_path) {
                    setDocUploadError('Carregue um arquivo ou informe um link de acesso.');
                    return;
                  }
                  saveDocumento(docDraft);
                }}
                disabled={docUploading}
                className={`px-5 py-2.5 rounded-xl font-bold text-sm text-white transition shadow-lg flex items-center gap-2 ${
                  docUploading ? 'bg-indigo-400 cursor-not-allowed' : 'bg-indigo-600 hover:bg-indigo-700 shadow-indigo-600/25'
                }`}
              >
                {docUploading ? (
                  <><span className="animate-spin inline-block w-4 h-4 border-2 border-white/30 border-t-white rounded-full"></span> Enviando...</>
                ) : (
                  <><Save size={16} /> {docDraft.dbId ? 'Salvar Alterações' : 'Adicionar Documento'}</>
                )}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );

  // Todas as abas disponíveis — filtradas por permissão abaixo
  const allNavItems = [
    { id: 'dashboard',         label: 'Dashboard Executivo',    icon: LayoutDashboard, section: null,                   adminOnly: true },
    { id: 'minhas_auditorias', label: 'Minhas Auditorias',      icon: ClipboardCheck,  section: 'Auditorias',           adminOnly: true },
    { id: 'checklist',         label: 'Lista de Verificação',   icon: ClipboardCheck,  section: null,                   adminOnly: true },
    { id: 'report',            label: 'Relatório RAI Final',    icon: FileText,        section: null,                   adminOnly: true },
    { id: 'rnc', label: 'Tratativa RNC (RAC)', icon: AlertTriangle, section: null, adminOnly: true, badge: rncs.filter((r) => r.status === 'Aberta').length },
    { id: 'gestao', label: 'Histórico de Auditorias', icon: FolderTree, section: 'Base de Conhecimento', adminOnly: true },
    { id: 'documentos', label: 'Documentos por Setor', icon: Folder, section: isAdmin ? null : null, adminOnly: false, badge: docsVencidos + docsProximos || 0, badgeColor: docsVencidos > 0 ? 'rose' : 'amber' },
    { id: 'calendario', label: 'Calendário de Auditorias', icon: Calendar, section: 'Programação', adminOnly: true }
  ];
  const navItems = allNavItems.filter((item) => isAdmin || !item.adminOnly);


  return (
    <div className="min-h-screen bg-[#F4F5F9] flex font-sans">
      {/* SIDEBAR */}
      <aside className="w-[280px] bg-[#0B1020] text-slate-300 flex flex-col shadow-2xl z-20 hidden md:flex print:hidden">
        <div className="p-6 border-b border-white/5 bg-gradient-to-br from-[#0B1020] to-[#10142B]">
          <KalenLogo variant="sidebar" />
          <div className="mt-3 flex items-center gap-2">
            <div className="w-5 h-px bg-indigo-600/60"></div>
            <p className="text-[10px] text-slate-500 uppercase tracking-widest font-bold">Sistema de Gestão de Auditoria</p>
          </div>
        </div>

        <div className="px-6 py-4 border-b border-white/5">
          <label className="text-[11px] font-bold text-slate-500 uppercase tracking-wider mb-2 block">Unidade Selecionada</label>
          <div className="relative">
            <select
              value={selectedBranch} onChange={(e) => setSelectedBranch(e.target.value)}
              className="w-full bg-white/5 border border-white/10 text-white rounded-xl p-3 pl-9 text-sm focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 appearance-none cursor-pointer"
            >
              {BRANCHES.map((branch) => <option key={branch} value={branch} className="text-black">{branch}</option>)}
            </select>
            <Building2 size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none" />
            <ChevronDown size={14} className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none" />
          </div>
        </div>

        <nav className="flex-1 mt-4 px-4 space-y-1 overflow-y-auto">
          {navItems.map((item) => (
            <React.Fragment key={item.id}>
              {item.section && (
                <div className="pt-4 pb-2">
                  <p className="px-4 text-[11px] font-bold text-slate-500 uppercase tracking-wider">{item.section}</p>
                </div>
              )}
              <button
                onClick={() => setActiveTab(item.id)}
                className={`w-full flex items-center gap-3 px-4 py-3 rounded-xl transition-all ${activeTab === item.id ? 'bg-indigo-600/15 text-indigo-300 font-bold' : 'hover:bg-white/5 hover:text-white font-medium text-slate-400'}`}
              >
                <item.icon size={19} className={activeTab === item.id ? 'text-indigo-400' : ''} />
                {item.label}
                {!!item.badge && (
                  <span className={`ml-auto text-white text-[10px] font-black px-2 py-0.5 rounded-full ${item.badgeColor === 'amber' ? 'bg-amber-500' : 'bg-rose-500'}`}>{item.badge}</span>
                )}
              </button>
            </React.Fragment>
          ))}
        </nav>

        <div className="p-4 m-4 space-y-2">
          {/* Info do usuário logado */}
          <div className="rounded-xl bg-white/5 border border-white/5 p-3">
            <div className="flex items-center gap-2.5">
              <div className="w-8 h-8 bg-indigo-600/30 text-indigo-300 rounded-full flex items-center justify-center font-black text-sm shrink-0">
                {currentUser.nome.charAt(0).toUpperCase()}
              </div>
              <div className="flex-1 min-w-0">
                <p className="text-xs font-bold text-slate-300 truncate">{currentUser.nome}</p>
                <p className="text-[10px] text-slate-500 truncate">{isAdmin ? 'Administrador' : currentUser.setor}</p>
              </div>
            </div>

            {/* Painel de gestão de auditores — só para admin */}
            {isAdmin && (
              <button
                onClick={() => setShowAdminPanel(true)}
                className="mt-2.5 w-full text-[10px] font-bold text-indigo-400 hover:text-indigo-300 border border-indigo-800/50 hover:border-indigo-600/50 rounded-lg py-1.5 transition flex items-center justify-center gap-1"
              >
                <ShieldCheck size={11} /> Gerenciar Auditores
              </button>
            )}

            <button
              onClick={onLogout}
              className="mt-2 w-full text-[10px] font-bold text-slate-500 hover:text-rose-400 border border-white/5 hover:border-rose-800/50 rounded-lg py-1.5 transition flex items-center justify-center gap-1"
            >
              <X size={10} /> Sair do Portal
            </button>
          </div>
          <p className="text-center text-[9px] text-slate-600">Kalenborn do Brasil · SGA v2.0 · ISO 9001:2015</p>
        </div>
      </aside>

      {/* ÁREA PRINCIPAL */}
      <div className="flex-1 flex flex-col min-w-0 h-screen overflow-hidden">

        <header className="h-16 bg-white border-b border-slate-200 shadow-sm flex items-center justify-between px-8 z-10 print:hidden">
          <div className="flex items-center gap-4 text-slate-500 font-medium text-sm">
            <span className="flex items-center gap-2 bg-slate-100 px-3 py-1.5 rounded-lg text-slate-700 font-bold">
              <Building2 size={16} className="text-indigo-600" /> {selectedBranch}
            </span>
            {!isAdmin && (
              <span className="bg-indigo-50 text-indigo-700 px-3 py-1.5 rounded-lg text-xs font-bold flex items-center gap-1.5">
                <Folder size={13} /> {userSetores[0]}
              </span>
            )}
            <span className="hidden lg:flex items-center gap-1.5 text-slate-400 text-xs">
              <Clock size={13} /> Atualizado em {formattedDate}
            </span>
            {isAdmin && (
              <span className={`flex items-center gap-1.5 text-xs font-bold transition-all ${
                autoSaveStatus === 'saved' ? 'text-emerald-600' :
                autoSaveStatus === 'unsaved' ? 'text-amber-500' :
                autoSaveStatus === 'saving' ? 'text-indigo-500 animate-pulse' :
                'text-rose-600'
              }`}>
                {autoSaveStatus === 'saved' && <><Check size={12} /> Salvo{lastSavedAt ? ` às ${lastSavedAt.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' })}` : ''}</>}
                {autoSaveStatus === 'unsaved' && <><AlertCircle size={12} /> Não salvo</>}
                {autoSaveStatus === 'saving' && <><Clock size={12} /> Salvando...</>}
                {autoSaveStatus === 'error' && <><AlertCircle size={12} /> Erro ao salvar</>}
              </span>
            )}
            {dbLoading && (
              <span className="flex items-center gap-1.5 text-amber-600 text-xs font-bold animate-pulse">
                <Clock size={12} /> Carregando dados do servidor...
              </span>
            )}
            {dbError && (
              <span className="flex items-center gap-1.5 text-rose-600 text-xs font-bold" title={dbError}>
                <AlertCircle size={12} /> Modo offline
              </span>
            )}
          </div>

          <div className="flex items-center gap-4">
            {isAdmin && (
              <div className="relative cursor-pointer">
                <Bell size={20} className="text-slate-500 hover:text-indigo-600 transition" />
                {rncs.filter((r) => r.status === 'Aberta').length > 0 && (
                  <div className="absolute -top-1 -right-1 w-2.5 h-2.5 bg-rose-500 rounded-full border-2 border-white"></div>
                )}
              </div>
            )}
            <div className="h-6 w-px bg-slate-200"></div>
            <div className="flex items-center gap-3">
              <div className="text-right hidden md:block">
                <p className="text-sm font-bold text-slate-800">{currentUser.nome}</p>
                <p className="text-[10px] text-slate-500 uppercase tracking-widest font-bold">
                  {isAdmin ? 'Administrador · SGQ' : currentUser.setor}
                </p>
              </div>
              <button
                onClick={onLogout}
                className="w-10 h-10 bg-indigo-100 text-indigo-700 rounded-full flex items-center justify-center font-black border-2 border-white shadow-sm hover:bg-rose-100 hover:text-rose-700 transition text-sm"
                title="Sair"
              >
                {currentUser.nome.charAt(0).toUpperCase()}
              </button>
            </div>
          </div>
        </header>

        <main className="flex-1 overflow-auto bg-[#F4F5F9] p-8 print:p-0 print:bg-white">
          <div className="max-w-7xl mx-auto print-area">
            {activeTab === 'dashboard'         && renderDashboard()}
            {activeTab === 'minhas_auditorias' && renderMinhasAuditorias()}
            {/* Antes dependia de "auditoriaAtiva" (só existe para RASCUNHOS em
                andamento) — então VER numa auditoria já finalizada carregava os
                dados certinho mas a tela continuava presa em "Minhas Auditorias".
                A condição certa é: existe checklist carregado pra mostrar? */}
            {/* Enquanto o VER ainda está buscando os dados no banco, mostra um
                carregando em vez de cair no fallback "Minhas Auditorias" — sem
                isso, clicar na aba rápido demais parecia que o checklist tinha
                sumido, quando na verdade a busca só ainda não tinha terminado. */}
            {activeTab === 'checklist' && (
              loadingAuditId ? renderCarregandoAuditoria() : (checklist.length > 0 ? renderChecklist() : renderMinhasAuditorias())
            )}
            {activeTab === 'report' && (
              loadingAuditId ? renderCarregandoAuditoria() : (checklist.length > 0 ? renderReport() : renderMinhasAuditorias())
            )}
            {activeTab === 'rnc' && renderRNC()}
            {activeTab === 'gestao' && renderGestao()}
            {activeTab === 'documentos' && renderDocumentos()}
            {activeTab === 'calendario' && renderCalendario()}
          </div>
        </main>
      </div>

      {/* MODAL DE GESTÃO DE AUDITORES — só para admin */}
      {showAdminPanel && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4 print:hidden" onClick={() => setShowAdminPanel(false)}>
          <div className="bg-white rounded-2xl shadow-2xl max-w-lg w-full p-6 max-h-[90vh] overflow-y-auto" onClick={(e) => e.stopPropagation()}>
            <div className="flex items-start justify-between mb-5">
              <div>
                <h3 className="text-lg font-black text-slate-900">Gerenciar Auditores</h3>
                <p className="text-xs text-slate-500 mt-1">Selecione as pessoas que terão acesso de administrador ao portal (mesmo nível que você)</p>
              </div>
              <button onClick={() => setShowAdminPanel(false)} className="text-slate-400 hover:text-slate-600"><X size={18} /></button>
            </div>

            <div className="space-y-2">
              {USUARIOS_BASE.filter((u) => u.email !== 'luciene.batista@kalenborn.com.br').map((u) => {
                const isPromovido = extraAdmins.includes(u.email.toLowerCase());
                return (
                  <div key={u.email} className="flex items-center justify-between p-3 rounded-xl border border-slate-100 hover:bg-slate-50 transition">
                    <div>
                      <p className="text-sm font-bold text-slate-800">{u.nome}</p>
                      <p className="text-xs text-slate-500">{u.email} · <span className="text-indigo-600">{u.setor}</span></p>
                    </div>
                    <button
                      onClick={() => toggleExtraAdmin(u.email, isPromovido)}
                      className={`px-3 py-1.5 rounded-lg text-xs font-bold border transition ${
                        isPromovido
                          ? 'bg-indigo-600 text-white border-indigo-600 hover:bg-rose-600 hover:border-rose-600'
                          : 'bg-white text-slate-500 border-slate-200 hover:bg-indigo-50 hover:border-indigo-300 hover:text-indigo-600'
                      }`}
                    >
                      {isPromovido ? 'Auditor ✓' : 'Promover'}
                    </button>
                  </div>
                );
              })}
            </div>

            {extraAdmins.length > 0 && (
              <div className="mt-4 bg-indigo-50 border border-indigo-100 rounded-xl p-3 text-xs text-indigo-700">
                <span className="font-bold">{extraAdmins.length} auditor(es) promovido(s)</span> — têm acesso completo ao portal enquanto esta configuração estiver ativa.
              </div>
            )}

            <button onClick={() => setShowAdminPanel(false)} className="mt-5 w-full py-2.5 rounded-xl font-bold text-sm bg-slate-100 text-slate-600 hover:bg-slate-200 transition">
              Fechar
            </button>
          </div>
        </div>
      )}

      {/* MODAL DE EVENTO — programar ou editar uma auditoria no calendário */}
      {eventDraft && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4 print:hidden" onClick={() => setEventDraft(null)}>
          <div className="bg-white rounded-2xl shadow-2xl max-w-md w-full p-6" onClick={(e) => e.stopPropagation()}>
            <div className="flex items-start justify-between mb-5">
              <div>
                <h3 className="text-lg font-black text-slate-900">{eventDraft.id ? 'Editar Auditoria Programada' : 'Programar Nova Auditoria'}</h3>
                <p className="text-xs text-slate-500 mt-1">{dateKeyToDate(eventDraft.dateKey).toLocaleDateString('pt-BR', { day: '2-digit', month: 'long', year: 'numeric', weekday: 'long' })}</p>
              </div>
              <button onClick={() => setEventDraft(null)} aria-label="Fechar" className="text-slate-400 hover:text-slate-600 shrink-0">
                <X size={18} />
              </button>
            </div>

            <div className="space-y-4">
              <div>
                <label className="text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5 block">Setor</label>
                <select
                  value={eventDraft.sector}
                  onChange={(e) => setEventDraft({ ...eventDraft, sector: e.target.value })}
                  className="w-full border border-slate-200 rounded-xl p-2.5 text-sm focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500"
                >
                  {SECTORS.map((s) => <option key={s} value={s}>{s}</option>)}
                </select>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5 block">Data</label>
                  <input
                    type="date"
                    value={eventDraft.dateKey}
                    onChange={(e) => setEventDraft({ ...eventDraft, dateKey: e.target.value })}
                    className="w-full border border-slate-200 rounded-xl p-2.5 text-sm focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500"
                  />
                </div>
                <div>
                  <label className="text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5 block">Horário</label>
                  <input
                    type="time"
                    value={eventDraft.time}
                    onChange={(e) => setEventDraft({ ...eventDraft, time: e.target.value })}
                    className="w-full border border-slate-200 rounded-xl p-2.5 text-sm focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500"
                  />
                </div>
              </div>

              <div>
                <label className="text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5 block">Duração (minutos)</label>
                <input
                  type="number"
                  min="15"
                  step="15"
                  value={eventDraft.durationMin}
                  onChange={(e) => setEventDraft({ ...eventDraft, durationMin: Number(e.target.value) || 60 })}
                  className="w-full border border-slate-200 rounded-xl p-2.5 text-sm focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500"
                />
              </div>

              <div>
                <label className="text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5 block">Auditor Líder</label>
                <select
                  value={eventDraft.auditor}
                  onChange={(e) => setEventDraft({ ...eventDraft, auditor: e.target.value })}
                  className="w-full border border-slate-200 rounded-xl p-2.5 text-sm focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500"
                >
                  {auditoresDb.map((a) => <option key={a.id} value={a.nome}>{a.nome}</option>)}
                </select>
              </div>

              <div>
                <label className="text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5 block">Observações (opcional)</label>
                <textarea
                  value={eventDraft.notes}
                  onChange={(e) => setEventDraft({ ...eventDraft, notes: e.target.value })}
                  rows="2"
                  placeholder="Ex.: focar em registros de calibração..."
                  className="w-full border border-slate-200 rounded-xl p-2.5 text-sm focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 resize-none"
                />
              </div>
            </div>

            <div className="flex justify-between items-center mt-6">
              {eventDraft.id ? (
                <button
                  onClick={() => {
                    const idToDelete = eventDraft.id;
                    setEventDraft(null);
                    requestConfirm({
                      title: 'Cancelar auditoria programada',
                      message: 'Tem certeza que deseja remover esta auditoria da programação?',
                      confirmLabel: 'Sim, remover',
                      danger: true,
                      onConfirm: () => deleteEvent(idToDelete)
                    });
                  }}
                  className="text-rose-600 hover:text-rose-700 font-bold text-sm flex items-center gap-1.5"
                >
                  <Trash2 size={15} /> Remover
                </button>
              ) : <div />}
              <div className="flex gap-3">
                <button onClick={() => setEventDraft(null)} className="px-4 py-2.5 rounded-xl font-bold text-sm text-slate-600 hover:bg-slate-100 transition">
                  Cancelar
                </button>
                <button
                  onClick={() => saveEvent(eventDraft)}
                  className="px-5 py-2.5 rounded-xl font-bold text-sm text-white bg-indigo-600 hover:bg-indigo-700 transition shadow-lg shadow-indigo-600/25 flex items-center gap-2"
                >
                  <Save size={16} /> {eventDraft.id ? 'Salvar Alterações' : 'Programar Auditoria'}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* MODAL DE E-MAIL — rascunho gerado automaticamente ao agendar/reagendar */}
      {emailDraft && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4 print:hidden" onClick={() => setEmailDraft(null)}>
          <div className="bg-white rounded-2xl shadow-2xl max-w-lg w-full p-6" onClick={(e) => e.stopPropagation()}>
            <div className="flex items-start gap-4 mb-2">
              <div className="p-2.5 rounded-xl shrink-0 bg-indigo-100 text-indigo-600">
                <Mail size={22} />
              </div>
              <div>
                <h3 className="text-lg font-black text-slate-900">Rascunho de notificação</h3>
                <p className="text-sm text-slate-500 mt-1">
                  Auditoria programada para o setor <span className="font-bold">{emailDraft.sector}</span>.
                </p>
              </div>
              <button onClick={() => setEmailDraft(null)} aria-label="Fechar" className="ml-auto text-slate-400 hover:text-slate-600 shrink-0">
                <X size={18} />
              </button>
            </div>

            <div className="bg-amber-50 border border-amber-200 text-amber-800 rounded-xl p-3 mb-4 flex items-start gap-2.5 text-xs">
              <AlertCircle size={15} className="mt-0.5 shrink-0" />
              <div>
                <span className="font-bold">Este e-mail ainda não foi enviado.</span> O sistema ainda não está conectado a um serviço de envio.
                Copie o texto abaixo e envie pelo seu cliente de e-mail (Outlook/Gmail) para notificar o setor de verdade.
              </div>
            </div>

            <div className="bg-slate-50 border border-slate-200 rounded-xl p-4 space-y-2">
              <div>
                <p className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Assunto</p>
                <p className="text-sm font-bold text-slate-800">{emailDraft.subject}</p>
              </div>
              <div>
                <p className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Mensagem</p>
                <pre className="text-sm text-slate-700 whitespace-pre-wrap font-sans">{emailDraft.body}</pre>
              </div>
            </div>

            <div className="flex justify-end gap-3 mt-6">
              <button onClick={() => setEmailDraft(null)} className="px-4 py-2.5 rounded-xl font-bold text-sm text-slate-600 hover:bg-slate-100 transition">
                Fechar
              </button>
              <button
                onClick={() => {
                  const fullText = `Assunto: ${emailDraft.subject}\n\n${emailDraft.body}`;
                  if (navigator.clipboard && navigator.clipboard.writeText) {
                    navigator.clipboard.writeText(fullText)
                      .then(() => setEmailCopied(true))
                      .catch(() => notify('Não foi possível copiar', 'Copie manualmente o texto acima.'));
                  } else {
                    notify('Não foi possível copiar', 'Seu navegador não permite copiar automaticamente. Selecione e copie o texto acima manualmente.');
                  }
                }}
                className={`px-5 py-2.5 rounded-xl font-bold text-sm text-white transition shadow-lg flex items-center gap-2 ${emailCopied ? 'bg-emerald-600 shadow-emerald-600/25' : 'bg-indigo-600 hover:bg-indigo-700 shadow-indigo-600/25'}`}
              >
                {emailCopied ? <Check size={16} /> : <Copy size={16} />} {emailCopied ? 'Copiado!' : 'Copiar e-mail'}
              </button>
            </div>

          </div>
        </div>
      )}

      {/* MODAL DE CONFIRMAÇÃO — usado antes de qualquer ação destrutiva ou irreversível */}
      {confirmAction && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4 print:hidden" onClick={closeConfirm}>
          <div className="bg-white rounded-2xl shadow-2xl max-w-md w-full p-6" onClick={(e) => e.stopPropagation()}>
            <div className="flex items-start gap-4 mb-4">
              <div className={`p-2.5 rounded-xl shrink-0 ${confirmAction.danger ? 'bg-rose-100 text-rose-600' : 'bg-indigo-100 text-indigo-600'}`}>
                <AlertCircle size={22} />
              </div>
              <div>
                <h3 className="text-lg font-black text-slate-900">{confirmAction.title}</h3>
                <p className="text-sm text-slate-500 mt-1.5 leading-relaxed">{confirmAction.message}</p>
              </div>
              <button onClick={closeConfirm} aria-label="Fechar" className="ml-auto text-slate-400 hover:text-slate-600 shrink-0">
                <X size={18} />
              </button>
            </div>
            <div className="flex justify-end gap-3 mt-6">
              <button onClick={closeConfirm} className="px-4 py-2.5 rounded-xl font-bold text-sm text-slate-600 hover:bg-slate-100 transition">
                Cancelar
              </button>
              <button
                onClick={handleConfirmedAction}
                className={`px-5 py-2.5 rounded-xl font-bold text-sm text-white transition shadow-lg ${confirmAction.danger ? 'bg-rose-600 hover:bg-rose-700 shadow-rose-600/25' : 'bg-indigo-600 hover:bg-indigo-700 shadow-indigo-600/25'}`}
              >
                {confirmAction.confirmLabel || 'Confirmar'}
              </button>
            </div>
          </div>
        </div>
      )}
      {/* Estilos globais em src/index.css */}
    </div>
  );
}
