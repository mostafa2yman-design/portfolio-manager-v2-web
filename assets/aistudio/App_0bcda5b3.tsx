import React, { useEffect, useMemo, useState } from 'react';
import {
  CalendarDays,
  List,
  Search,
  Plus,
  ChevronLeft,
  ChevronRight,
  UserRound,
  CheckCircle2,
  CircleAlert,
  Download,
  X,
  Save,
  CalendarCheck,
  Upload,
  FileSpreadsheet,
  Trash2,
  MessageCircle,
  Phone,
  AlertTriangle,
  Clock,
  Edit3
} from 'lucide-react';
import * as XLSX from 'xlsx';

export type Status = 'جديدة' | 'قيد التنفيذ' | 'تمت' | 'مؤجلة';

export type Customer = {
  id: string;
  name: string;
  phone: string;
  installment: number | null;
  daysLate: number | null;
  dueDate: string;
  totalOverdue: number | null;
  loanBalance: number | null;
  currentAccountBalance: number | null;
  currentAccount: string;
  cif: string;
  address: string;
  notes: string;
};

export type Task = {
  id: string;
  customerId: string;
  date: string;
  time: string;
  kind: string;
  notes: string;
  status: Status;
};

export const CUSTOMER_COLUMNS = [
  'رقم التليفون',
  'اسم العميل',
  'قيمة القسط',
  'عدد أيام التأخير',
  'تاريخ الاستحقاق',
  'إجمالي متأخرات',
  'رصيد القرض',
  'رصيد الحساب الجاري',
  'الحساب الجاري',
  'رقم الـ CIF',
  'العنوان'
];

const KEY = 'customer-followup-v1';

const fmt = (d: string) =>
  d
    ? new Date(d + 'T12:00:00').toLocaleDateString('ar-EG', {
        weekday: 'long',
        day: 'numeric',
        month: 'long',
        year: 'numeric'
      })
    : '—';

const monthLabel = (y: number, m: number) =>
  new Date(y, m, 1).toLocaleDateString('ar-EG', {
    month: 'long',
    year: 'numeric'
  });

const statusClass = (s: Status) =>
  s === 'تمت' ? 'done' : s === 'مؤجلة' ? 'postponed' : s === 'قيد التنفيذ' ? 'progress' : 'new';

const uid = (prefix: string) =>
  prefix + '-' + Math.random().toString(36).slice(2, 8).toUpperCase();

const numeric = (v: any): number | null => {
  if (v === null || v === undefined || String(v).trim() === '') return null;
  const n = Number(String(v).replace(/[,،]/g, '').trim());
  return Number.isFinite(n) ? n : null;
};

const normalize = (v: any) =>
  String(v ?? '')
    .trim()
    .toLowerCase()
    .replace(/[\sـ]/g, '');

const getCleanPhone = (phone?: string) => {
  if (!phone) return '';
  let p = phone.replace(/[^0-9]/g, '');
  if (p.startsWith('01') && p.length === 11) {
    p = '2' + p;
  }
  return p;
};

const localToday = () => {
  const today = new Date();
  const year = today.getFullYear();
  const month = String(today.getMonth() + 1).padStart(2, '0');
  const day = String(today.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
};

export default function App() {
  const [customers, setCustomers] = useState<Customer[]>(() => {
    try {
      const stored = JSON.parse(localStorage.getItem(KEY + '-customers') || '[]');
      return Array.isArray(stored) ? stored : [];
    } catch {
      return [];
    }
  });

  const [tasks, setTasks] = useState<Task[]>(() => {
    try {
      const stored = JSON.parse(localStorage.getItem(KEY + '-tasks') || '[]');
      return Array.isArray(stored) ? stored : [];
    } catch {
      return [];
    }
  });

  const [view, setView] = useState<'calendar' | 'list' | 'customer'>('calendar');
  const [month, setMonth] = useState(() => {
    const today = new Date();
    return new Date(today.getFullYear(), today.getMonth(), 1);
  });
  const [selectedDate, setSelectedDate] = useState(localToday);
  const [selectedCustomer, setSelectedCustomer] = useState('');
  const [search, setSearch] = useState('');

  // Modals
  const [taskModal, setTaskModal] = useState(false);
  const [editing, setEditing] = useState<Task | null>(null);
  const [taskForm, setTaskForm] = useState({
    customerId: '',
    date: localToday(),
    time: '10:00',
    kind: 'اتصال هاتفي',
    notes: '',
    status: 'جديدة' as Status
  });

  const [customerModal, setCustomerModal] = useState(false);
  const [editingCustomer, setEditingCustomer] = useState<Customer | null>(null);
  const [newCustomer, setNewCustomer] = useState<Record<string, string>>({});

  const [importModal, setImportModal] = useState(false);
  const [importRows, setImportRows] = useState<Customer[] | null>(null);
  const [importError, setImportError] = useState('');
  const [importFileName, setImportFileName] = useState('');

  // Confirmation Modal
  const [deleteConfirm, setDeleteConfirm] = useState<{
    type: 'task' | 'customer' | 'clear_all';
    id?: string;
    title: string;
    message: string;
  } | null>(null);

  useEffect(() => {
    try {
      localStorage.setItem(KEY + '-customers', JSON.stringify(customers));
    } catch (e) {
      console.error(e);
    }
  }, [customers]);

  useEffect(() => {
    try {
      localStorage.setItem(KEY + '-tasks', JSON.stringify(tasks));
    } catch (e) {
      console.error(e);
    }
  }, [tasks]);

  const customer = (id: string) => customers.find((c) => c.id === id);

  useEffect(() => {
    if (customers.length > 0 && !customer(selectedCustomer)) {
      setSelectedCustomer(customers[0].id);
    }
  }, [customers, selectedCustomer]);

  const filtered = useMemo(
    () =>
      tasks
        .filter((t) => {
          const c = customer(t.customerId);
          const q = search.trim().toLowerCase();
          return (
            !q ||
            c?.name.toLowerCase().includes(q) ||
            c?.phone.includes(q) ||
            c?.cif.includes(q) ||
            t.notes.toLowerCase().includes(q) ||
            t.kind.toLowerCase().includes(q)
          );
        })
        .sort((a, b) => (a.date + a.time).localeCompare(b.date + b.time)),
    [tasks, customers, search]
  );

  const monthTasks = filtered.filter((t) =>
    t.date.startsWith(`${month.getFullYear()}-${String(month.getMonth() + 1).padStart(2, '0')}`)
  );
  const selectedTasks = filtered.filter((t) => t.date === selectedDate);
  const customerTasks = filtered
    .filter((t) => t.customerId === selectedCustomer)
    .sort((a, b) => (b.date + b.time).localeCompare(a.date + a.time));
  const done = tasks.filter((t) => t.status === 'تمت').length;

  function openAdd(date = selectedDate) {
    if (customers.length === 0) {
      alert('يرجى إضافة أو استيراد عميل أولاً قبل إنشاء متابعة.');
      setEditingCustomer(null);
      setNewCustomer({});
      setCustomerModal(true);
      return;
    }
    setEditing(null);
    setTaskForm({
      customerId: selectedCustomer || customers[0]?.id || '',
      date,
      time: '10:00',
      kind: 'اتصال هاتفي',
      notes: '',
      status: 'جديدة'
    });
    setTaskModal(true);
  }

  function openEdit(t: Task) {
    setEditing(t);
    setTaskForm({ ...t });
    setTaskModal(true);
  }

  function saveTask(e: React.FormEvent) {
    e.preventDefault();
    const t: Task = { ...taskForm, id: editing?.id || uid('T') };
    setTasks((old) => (editing ? old.map((x) => (x.id === t.id ? t : x)) : [...old, t]));
    setSelectedDate(t.date);
    const [y, m] = t.date.split('-').map(Number);
    if (!isNaN(y) && !isNaN(m)) {
      setMonth(new Date(y, m - 1, 1));
    }
    setTaskModal(false);
  }

  function deleteTask(id: string) {
    setTasks((old) => old.filter((t) => t.id !== id));
    setDeleteConfirm(null);
  }

  function openAddCustomer() {
    setEditingCustomer(null);
    setNewCustomer({});
    setCustomerModal(true);
  }

  function openEditCustomer(c: Customer) {
    setEditingCustomer(c);
    setNewCustomer({
      name: c.name || '',
      phone: c.phone || '',
      installment: c.installment !== null ? String(c.installment) : '',
      daysLate: c.daysLate !== null ? String(c.daysLate) : '',
      dueDate: c.dueDate || '',
      totalOverdue: c.totalOverdue !== null ? String(c.totalOverdue) : '',
      loanBalance: c.loanBalance !== null ? String(c.loanBalance) : '',
      currentAccountBalance: c.currentAccountBalance !== null ? String(c.currentAccountBalance) : '',
      currentAccount: c.currentAccount || '',
      cif: c.cif || '',
      address: c.address || '',
      notes: c.notes || ''
    });
    setCustomerModal(true);
  }

  function saveCustomer(e: React.FormEvent) {
    e.preventDefault();
    if (editingCustomer) {
      const updated: Customer = {
        ...editingCustomer,
        name: newCustomer.name || '',
        phone: newCustomer.phone || '',
        installment: numeric(newCustomer.installment),
        daysLate: numeric(newCustomer.daysLate),
        dueDate: newCustomer.dueDate || '',
        totalOverdue: numeric(newCustomer.totalOverdue),
        loanBalance: numeric(newCustomer.loanBalance),
        currentAccountBalance: numeric(newCustomer.currentAccountBalance),
        currentAccount: newCustomer.currentAccount || '',
        cif: newCustomer.cif || '',
        address: newCustomer.address || '',
        notes: newCustomer.notes || ''
      };
      setCustomers((old) => old.map((c) => (c.id === updated.id ? updated : c)));
      setSelectedCustomer(updated.id);
    } else {
      const c: Customer = {
        id: uid('C'),
        name: newCustomer.name || '',
        phone: newCustomer.phone || '',
        installment: numeric(newCustomer.installment),
        daysLate: numeric(newCustomer.daysLate),
        dueDate: newCustomer.dueDate || '',
        totalOverdue: numeric(newCustomer.totalOverdue),
        loanBalance: numeric(newCustomer.loanBalance),
        currentAccountBalance: numeric(newCustomer.currentAccountBalance),
        currentAccount: newCustomer.currentAccount || '',
        cif: newCustomer.cif || '',
        address: newCustomer.address || '',
        notes: newCustomer.notes || ''
      };
      setCustomers((old) => [...old, c]);
      setSelectedCustomer(c.id);
      setTaskForm((f) => ({ ...f, customerId: c.id }));
    }
    setCustomerModal(false);
    setNewCustomer({});
  }

  function deleteCustomer(id: string) {
    setCustomers((old) => old.filter((c) => c.id !== id));
    setTasks((old) => old.filter((t) => t.customerId !== id));
    if (selectedCustomer === id) {
      const remaining = customers.filter((c) => c.id !== id);
      setSelectedCustomer(remaining.length > 0 ? remaining[0].id : '');
    }
    setDeleteConfirm(null);
  }

  function parseImportFile(file: File) {
    setImportError('');
    setImportRows(null);
    setImportFileName(file.name);
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const wb = XLSX.read(reader.result, { type: 'array', cellDates: true });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json<any[]>(sheet, { header: 1, defval: '' });
        if (rows.length < 2) throw new Error('الملف لا يحتوي على صفوف بيانات.');
        const headers = rows[0].map(normalize);
        const indices = CUSTOMER_COLUMNS.map((h) => headers.indexOf(normalize(h)));
        if (indices.some((i) => i < 0)) {
          const missing = CUSTOMER_COLUMNS.filter((_, idx) => indices[idx] < 0);
          throw new Error(`الأعمدة غير مطابقة. تأكد من وجود الأعمدة الـ 11 المطلوبة. النواقص: ${missing.join('، ')}`);
        }
        const parsed = rows
          .slice(1)
          .filter((r: any[]) => r.some((v: any) => String(v ?? '').trim() !== ''))
          .map((r: any[]) => {
            const v = (i: number) => r[indices[i]] ?? '';
            const dateVal = v(4);
            let dueDate = '';
            if (dateVal instanceof Date) dueDate = dateVal.toISOString().slice(0, 10);
            else if (typeof dateVal === 'number') {
              const d = XLSX.SSF.parse_date_code(dateVal);
              if (d) dueDate = `${d.y}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}`;
            } else if (String(dateVal).trim()) {
              const raw = String(dateVal).trim();
              const d = new Date(raw);
              dueDate = Number.isNaN(d.getTime()) ? raw : d.toISOString().slice(0, 10);
            }
            return {
              id: uid('C'),
              phone: String(v(0)).trim(),
              name: String(v(1)).trim(),
              installment: numeric(v(2)),
              daysLate: numeric(v(3)),
              dueDate,
              totalOverdue: numeric(v(5)),
              loanBalance: numeric(v(6)),
              currentAccountBalance: numeric(v(7)),
              currentAccount: String(v(8)).trim(),
              cif: String(v(9)).trim(),
              address: String(v(10)).trim(),
              notes: ''
            } as Customer;
          });
        if (!parsed.length) throw new Error('لم يتم العثور على صفوف عملاء.');
        if (parsed.some((c) => !c.name || (!c.cif && !c.phone)))
          throw new Error('يوجد صف بدون اسم العميل أو بدون رقم CIF ورقم تليفون.');
        setImportRows(parsed);
      } catch (err: any) {
        setImportError(err.message || 'تعذر قراءة الملف.');
      }
    };
    reader.onerror = () => setImportError('تعذر قراءة الملف.');
    reader.readAsArrayBuffer(file);
  }

  function confirmImport() {
    if (!importRows) return;
    setCustomers((old) => {
      const next = [...old];
      for (const incoming of importRows) {
        const idx = next.findIndex(
          (c) =>
            (incoming.cif && c.cif === incoming.cif) ||
            (!incoming.cif && incoming.phone && c.phone === incoming.phone)
        );
        if (idx >= 0)
          next[idx] = {
            ...next[idx],
            ...incoming,
            id: next[idx].id,
            notes: next[idx].notes || incoming.notes
          };
        else next.push(incoming);
      }
      return next;
    });
    if (importRows.length > 0) {
      setSelectedCustomer(importRows[0].id);
    }
    setImportModal(false);
    setImportRows(null);
    setImportFileName('');
    setImportError('');
  }

  function exportData() {
    const blob = new Blob([JSON.stringify({ customers, tasks }, null, 2)], {
      type: 'application/json'
    });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'customer-followups-backup.json';
    a.click();
    URL.revokeObjectURL(a.href);
  }

  function changeStatus(id: string, status: Status) {
    setTasks((old) => old.map((t) => (t.id === id ? { ...t, status } : t)));
  }

  function clearAllData() {
    setCustomers([]);
    setTasks([]);
    setSelectedCustomer('');
    try {
      localStorage.removeItem(KEY + '-customers');
      localStorage.removeItem(KEY + '-tasks');
    } catch (e) {
      console.error(e);
    }
    setDeleteConfirm(null);
  }

  const days = Array.from(
    { length: new Date(month.getFullYear(), month.getMonth() + 1, 0).getDate() },
    (_, i) => i + 1
  );
  const first = (new Date(month.getFullYear(), month.getMonth(), 1).getDay() + 1) % 7;

  const activeCustomer = customer(selectedCustomer);

  const customerFields: [string, any][] = [
    ['قيمة القسط', activeCustomer?.installment],
    ['عدد أيام التأخير', activeCustomer?.daysLate],
    ['تاريخ الاستحقاق', activeCustomer?.dueDate],
    ['إجمالي متأخرات', activeCustomer?.totalOverdue],
    ['رصيد القرض', activeCustomer?.loanBalance],
    ['رصيد الحساب الجاري', activeCustomer?.currentAccountBalance],
    ['الحساب الجاري', activeCustomer?.currentAccount],
    ['رقم الـ CIF', activeCustomer?.cif],
    ['العنوان', activeCustomer?.address]
  ];

  return (
    <div className="app">
      {/* Sidebar */}
      <aside className="sidebar">
        <div className="brand">
          <div className="brandmark">
            <CalendarCheck />
          </div>
          <div>
            <b>متابعة العملاء</b>
            <small>إدارة المتابعات والملفات</small>
          </div>
        </div>

        <button
          className={view === 'calendar' ? 'nav active' : 'nav'}
          onClick={() => setView('calendar')}
        >
          <CalendarDays /> التقويم الشهري
        </button>

        <button
          className={view === 'list' ? 'nav active' : 'nav'}
          onClick={() => setView('list')}
        >
          <List /> الجدول اليومي
        </button>

        <button
          className={view === 'customer' ? 'nav active' : 'nav'}
          onClick={() => setView('customer')}
        >
          <UserRound /> سجل العملاء ({customers.length})
        </button>

        <div className="sidebottom">
          <div className="sidehint">
            البيانات محفوظة محليًا في هذا المتصفح. يمكنك استيراد ملفات Excel وتصدير نسخ احتياطية.
          </div>

          <button
            className="nav"
            onClick={() => {
              setImportModal(true);
              setImportRows(null);
              setImportError('');
            }}
          >
            <Upload /> استيراد العملاء من Excel
          </button>

          <button className="nav" onClick={exportData}>
            <Download /> تصدير نسخة احتياطية
          </button>

          <button
            className="nav danger font-bold"
            onClick={() =>
              setDeleteConfirm({
                type: 'clear_all',
                title: 'حذف جميع بيانات البرنامج نهائيًا؟',
                message: `تحذير: سيتم حذف كافة العملاء (${customers.length} عميل) وكافة المهام والمتابعات (${tasks.length} مهمة) نهائيًا.`
              })
            }
          >
            <Trash2 /> مسح وحذف كل البيانات
          </button>
        </div>
      </aside>

      {/* Main Content */}
      <main className="main">
        <header className="topbar">
          <div>
            <div className="eyebrow">إدارة علاقات العملاء والمتابعات البنكية والائتمانية</div>
            <h1>
              {view === 'calendar'
                ? 'التقويم والمتابعات'
                : view === 'list'
                ? 'الجدول اليومي'
                : 'تفاصيل وملف العميل'}
            </h1>
          </div>

          <div className="topactions">
            <label className="search">
              <Search size={17} />
              <input
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                placeholder="ابحث بالاسم أو التليفون أو CIF..."
              />
              {search && (
                <button
                  onClick={() => setSearch('')}
                  className="text-gray-400 hover:text-gray-600 border-0 bg-transparent p-0"
                >
                  <X size={15} />
                </button>
              )}
            </label>

            <button className="secondary" onClick={openAddCustomer}>
              <UserRound size={17} /> عميل جديد
            </button>

            <button className="primary" onClick={() => openAdd()}>
              <Plus size={18} /> إضافة متابعة
            </button>
          </div>
        </header>

        {/* CALENDAR VIEW */}
        {view === 'calendar' && (
          <>
            <section className="stats">
              <Stat icon={<CalendarDays />} label="إجمالي المهام" value={tasks.length} />
              <Stat icon={<CheckCircle2 />} label="تمت" value={done} />
              <Stat icon={<CircleAlert />} label="متبقية" value={tasks.length - done} />
            </section>

            <section className="panel">
              <div className="panelhead">
                <div>
                  <h2>{monthLabel(month.getFullYear(), month.getMonth())}</h2>
                  <p>اختار يومًا لعرض المتابعات أو إضافة مهمة جديدة</p>
                </div>

                <div className="switches">
                  <button onClick={() => setView('calendar')} className="selected">
                    <CalendarDays size={16} /> تقويم
                  </button>
                  <button onClick={() => setView('list')}>
                    <List size={16} /> جدول
                  </button>
                  <button
                    className="iconbtn"
                    title="الشهر السابق"
                    onClick={() => setMonth(new Date(month.getFullYear(), month.getMonth() - 1, 1))}
                  >
                    <ChevronRight />
                  </button>
                  <button
                    className="iconbtn"
                    title="الشهر التالي"
                    onClick={() => setMonth(new Date(month.getFullYear(), month.getMonth() + 1, 1))}
                  >
                    <ChevronLeft />
                  </button>
                </div>
              </div>

              <div className="calendar">
                <div className="weekdays">
                  {['السبت', 'الأحد', 'الاثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة'].map(
                    (x) => (
                      <div key={x}>{x}</div>
                    )
                  )}
                </div>

                <div className="daygrid">
                  {Array.from({ length: first }, (_, i) => (
                    <div key={'blank' + i} className="blank" />
                  ))}
                  {days.map((d) => {
                    const date = `${month.getFullYear()}-${String(month.getMonth() + 1).padStart(
                      2,
                      '0'
                    )}-${String(d).padStart(2, '0')}`;
                    const ts = monthTasks.filter((t) => t.date === date);
                    return (
                      <button
                        key={date}
                        className={'day ' + (date === selectedDate ? 'chosen' : '')}
                        onClick={() => setSelectedDate(date)}
                      >
                        <span className="daynum">{d}</span>
                        {ts.length > 0 && (
                          <span className="taskcount">
                            {ts.length} {ts.length === 1 ? 'مهمة' : 'مهام'}
                          </span>
                        )}
                        {ts.slice(0, 2).map((t) => (
                          <span className={'daytask ' + statusClass(t.status)} key={t.id}>
                            {customer(t.customerId)?.name || 'عميل'} · {t.kind}
                          </span>
                        ))}
                        {ts.length > 2 && <span className="more">+{ts.length - 2} أخرى</span>}
                      </button>
                    );
                  })}
                </div>
              </div>
            </section>

            <section className="panel">
              <div className="panelhead">
                <div>
                  <h2>{fmt(selectedDate)}</h2>
                  <p>{selectedTasks.length} متابعة في هذا اليوم</p>
                </div>
                <button className="secondary" onClick={() => openAdd(selectedDate)}>
                  <Plus size={17} /> أضف مهمة لهذا اليوم
                </button>
              </div>

              <TaskTable
                tasks={selectedTasks}
                customer={customer}
                onEdit={openEdit}
                onDelete={(id) =>
                  setDeleteConfirm({
                    type: 'task',
                    id,
                    title: 'حذف المتابعة؟',
                    message: 'هل أنت متأكد من حذف هذه المتابعة نهائيًا؟'
                  })
                }
                onStatus={changeStatus}
                onCustomer={(id) => {
                  setSelectedCustomer(id);
                  setView('customer');
                }}
              />
            </section>
          </>
        )}

        {/* LIST VIEW */}
        {view === 'list' && (
          <section className="panel">
            <div className="panelhead">
              <div>
                <h2>كل المتابعات</h2>
                <p>مرتبة حسب التاريخ والساعة</p>
              </div>
              <button className="secondary" onClick={() => openAdd()}>
                <Plus size={17} /> إضافة مهمة
              </button>
            </div>

            <TaskTable
              tasks={filtered}
              customer={customer}
              onEdit={openEdit}
              onDelete={(id) =>
                setDeleteConfirm({
                  type: 'task',
                  id,
                  title: 'حذف المتابعة؟',
                  message: 'هل أنت متأكد من حذف هذه المتابعة نهائيًا؟'
                })
              }
              onStatus={changeStatus}
              onCustomer={(id) => {
                setSelectedCustomer(id);
                setView('customer');
              }}
            />
          </section>
        )}

        {/* CUSTOMER PROFILE VIEW */}
        {view === 'customer' && (
          <>
            {activeCustomer ? (
              <>
                <section className="panel">
                  <div className="customerhero">
                    <div className="avatar">
                      <UserRound />
                    </div>
                    <div className="customerinfo">
                      <h2>{activeCustomer.name}</h2>
                      <p>
                        {activeCustomer.phone} {activeCustomer.cif ? `· CIF: ${activeCustomer.cif}` : ''}
                      </p>
                      <small>{activeCustomer.address || 'العنوان غير مسجل'}</small>
                      {activeCustomer.notes && (
                        <div className="text-xs text-gray-500 mt-1">
                          <b>ملاحظات:</b> {activeCustomer.notes}
                        </div>
                      )}
                    </div>

                    <div className="flex items-center gap-2 flex-wrap">
                      {activeCustomer.phone && (
                        <>
                          <a
                            href={`https://wa.me/${getCleanPhone(activeCustomer.phone)}`}
                            target="_blank"
                            rel="noreferrer"
                            className="primary text-xs !bg-[#25D366] hover:!bg-[#1faa53] !no-underline"
                            title="واتساب"
                          >
                            <MessageCircle size={15} /> واتساب
                          </a>
                          <a
                            href={`tel:${activeCustomer.phone}`}
                            className="secondary text-xs !no-underline"
                            title="اتصال"
                          >
                            <Phone size={15} /> اتصال
                          </a>
                        </>
                      )}
                      <button
                        className="secondary text-xs"
                        onClick={() => openEditCustomer(activeCustomer)}
                      >
                        <Edit3 size={14} /> تعديل البيانات
                      </button>
                      <button className="primary" onClick={() => openAdd()}>
                        <Plus size={18} /> متابعة جديدة
                      </button>
                    </div>
                  </div>

                  {/* 11 Customer columns data grid */}
                  <div className="customerdata">
                    {customerFields.map(([label, value]) => (
                      <div className="customerfield" key={label}>
                        <small>{label}</small>
                        <b>{value === null || value === undefined || value === '' ? '—' : String(value)}</b>
                      </div>
                    ))}
                  </div>

                  <div className="customerselect">
                    <label>اختيار العميل</label>
                    <select
                      value={selectedCustomer}
                      onChange={(e) => setSelectedCustomer(e.target.value)}
                    >
                      {customers.map((c) => (
                        <option value={c.id} key={c.id}>
                          {c.name} — {c.phone} {c.cif ? `(CIF: ${c.cif})` : ''}
                        </option>
                      ))}
                    </select>
                    <button className="secondary" onClick={openAddCustomer}>
                      <Plus size={16} /> عميل جديد
                    </button>
                    <button
                      className="dangerbtn text-xs"
                      onClick={() =>
                        setDeleteConfirm({
                          type: 'customer',
                          id: activeCustomer.id,
                          title: `حذف العميل "${activeCustomer.name}"؟`,
                          message:
                            'سيؤدي حذف العميل إلى حذف كافة المتابعات والمهام المرتبطة به نهائيًا.'
                        })
                      }
                    >
                      <Trash2 size={14} /> حذف العميل
                    </button>
                  </div>
                </section>

                <section className="panel">
                  <div className="panelhead">
                    <div>
                      <h2>سجل المتابعات</h2>
                      <p>{customerTasks.length} متابعة محفوظة لهذا العميل</p>
                    </div>
                    <button className="secondary" onClick={() => openAdd()}>
                      <Plus size={16} /> إضافة متابعة للعميل
                    </button>
                  </div>

                  {customerTasks.length === 0 ? (
                    <div className="empty">
                      <CalendarDays />
                      <b>لا توجد متابعات مسجلة لهذا العميل حتى الآن</b>
                      <span>يمكنك تسجيل أول متابعة له بالضغط على الزر بالأعلى.</span>
                    </div>
                  ) : (
                    <div className="timeline">
                      {customerTasks.map((t) => (
                        <div className="timelineitem" key={t.id}>
                          <div className="timelinepin" />
                          <div className="timelinebody">
                            <div className="timelinehead">
                              <b>{fmt(t.date)}</b>
                              <StatusBadge status={t.status} />
                            </div>
                            <div className="tasktitle">
                              {t.kind} · {t.time}
                            </div>
                            <p>{t.notes || 'لا توجد ملاحظات'}</p>
                            <div className="flex items-center gap-3 mt-2">
                              <button className="textbutton" onClick={() => openEdit(t)}>
                                تعديل المتابعة
                              </button>
                              <button
                                className="dangertextbtn"
                                onClick={() =>
                                  setDeleteConfirm({
                                    type: 'task',
                                    id: t.id,
                                    title: 'حذف المتابعة؟',
                                    message: 'هل تريد حذف هذه المتابعة من سجل العميل؟'
                                  })
                                }
                              >
                                حذف
                              </button>
                            </div>
                          </div>
                        </div>
                      ))}
                    </div>
                  )}
                </section>
              </>
            ) : (
              <div className="panel empty">
                <UserRound size={32} />
                <b>لا يوجد عملاء مضافون حاليًا</b>
                <p className="text-xs text-gray-500">
                  يمكنك استيراد ملف Excel أو إضافة العملاء يدويًا.
                </p>
                <div className="flex items-center gap-2 mt-3">
                  <button className="primary" onClick={openAddCustomer}>
                    <Plus size={17} /> إضافة عميل جديد
                  </button>
                  <button
                    className="secondary"
                    onClick={() => {
                      setImportModal(true);
                      setImportRows(null);
                      setImportError('');
                    }}
                  >
                    <Upload size={17} /> استيراد Excel
                  </button>
                </div>
              </div>
            )}
          </>
        )}
      </main>

      {/* TASK MODAL */}
      {taskModal && (
        <div
          className="overlay"
          onMouseDown={(e) => {
            if (e.target === e.currentTarget) setTaskModal(false);
          }}
        >
          <form className="modal" onSubmit={saveTask}>
            <div className="modalhead">
              <div>
                <h2>{editing ? 'تعديل المتابعة' : 'إضافة متابعة جديدة'}</h2>
                <p>تسجيل متابعة مستقلة في سجل العميل</p>
              </div>
              <button type="button" className="iconbtn" onClick={() => setTaskModal(false)}>
                <X />
              </button>
            </div>

            <div className="formgrid">
              <div className="field full">
                <label>العميل</label>
                <div className="inlinefield">
                  <select
                    required
                    value={taskForm.customerId}
                    onChange={(e) => setTaskForm({ ...taskForm, customerId: e.target.value })}
                  >
                    {customers.map((c) => (
                      <option key={c.id} value={c.id}>
                        {c.name} — {c.phone} {c.cif ? `(CIF: ${c.cif})` : ''}
                      </option>
                    ))}
                  </select>
                  <button type="button" className="secondary" onClick={openAddCustomer}>
                    عميل جديد
                  </button>
                </div>
              </div>

              <div className="field">
                <label>تاريخ المتابعة</label>
                <input
                  type="date"
                  required
                  value={taskForm.date}
                  onChange={(e) => setTaskForm({ ...taskForm, date: e.target.value })}
                />
              </div>

              <div className="field">
                <label>الساعة</label>
                <input
                  type="time"
                  required
                  value={taskForm.time}
                  onChange={(e) => setTaskForm({ ...taskForm, time: e.target.value })}
                />
              </div>

              <div className="field">
                <label>نوع المهمة</label>
                <select
                  value={taskForm.kind}
                  onChange={(e) => setTaskForm({ ...taskForm, kind: e.target.value })}
                >
                  {[
                    'اتصال هاتفي',
                    'زيارة العميل',
                    'استكمال مستندات',
                    'متابعة طلب التمويل',
                    'تحصيل',
                    'عرض أسعار',
                    'أخرى'
                  ].map((x) => (
                    <option key={x}>{x}</option>
                  ))}
                </select>
              </div>

              <div className="field">
                <label>حالة المهمة</label>
                <select
                  value={taskForm.status}
                  onChange={(e) =>
                    setTaskForm({ ...taskForm, status: e.target.value as Status })
                  }
                >
                  {['جديدة', 'قيد التنفيذ', 'تمت', 'مؤجلة'].map((x) => (
                    <option key={x}>{x}</option>
                  ))}
                </select>
              </div>

              <div className="field full">
                <label>ملاحظات المتابعة</label>
                <textarea
                  rows={4}
                  value={taskForm.notes}
                  onChange={(e) => setTaskForm({ ...taskForm, notes: e.target.value })}
                  placeholder="اكتب المطلوب أو نتيجة المتابعة..."
                />
              </div>
            </div>

            <div className="modalfoot">
              <button type="button" className="secondary" onClick={() => setTaskModal(false)}>
                إلغاء
              </button>
              <button className="primary" type="submit">
                <Save size={17} /> حفظ المتابعة
              </button>
            </div>
          </form>
        </div>
      )}

      {/* CUSTOMER MODAL (11 COLUMNS) */}
      {customerModal && (
        <div
          className="overlay"
          onMouseDown={(e) => {
            if (e.target === e.currentTarget) setCustomerModal(false);
          }}
        >
          <form className="modal customereditmodal" onSubmit={saveCustomer}>
            <div className="modalhead">
              <div>
                <h2>{editingCustomer ? 'تعديل بيانات العميل' : 'إضافة عميل جديد'}</h2>
                <p>إدخال بيانات العميل والحسابات الائتمانية والبنكية</p>
              </div>
              <button type="button" className="iconbtn" onClick={() => setCustomerModal(false)}>
                <X />
              </button>
            </div>

            <div className="formgrid">
              {[
                ['phone', 'رقم التليفون', 'tel'],
                ['name', 'اسم العميل', 'text'],
                ['installment', 'قيمة القسط', 'number'],
                ['daysLate', 'عدد أيام التأخير', 'number'],
                ['dueDate', 'تاريخ الاستحقاق', 'date'],
                ['totalOverdue', 'إجمالي متأخرات', 'number'],
                ['loanBalance', 'رصيد القرض', 'number'],
                ['currentAccountBalance', 'رصيد الحساب الجاري', 'number'],
                ['currentAccount', 'الحساب الجاري', 'text'],
                ['cif', 'رقم الـ CIF', 'text'],
                ['address', 'العنوان', 'text']
              ].map(([key, label, inputType]) => (
                <div className="field" key={key}>
                  <label>{label}</label>
                  <input
                    required={key === 'name' || key === 'phone'}
                    type={inputType}
                    value={newCustomer[key] || ''}
                    onChange={(e) => setNewCustomer({ ...newCustomer, [key]: e.target.value })}
                  />
                </div>
              ))}

              <div className="field full">
                <label>ملاحظات إضافية حول العميل</label>
                <textarea
                  rows={2}
                  value={newCustomer.notes || ''}
                  onChange={(e) => setNewCustomer({ ...newCustomer, notes: e.target.value })}
                  placeholder="أي ملاحظات هامة..."
                />
              </div>
            </div>

            <div className="modalfoot">
              <button type="button" className="secondary" onClick={() => setCustomerModal(false)}>
                إلغاء
              </button>
              <button className="primary" type="submit">
                <Save size={17} /> حفظ العميل
              </button>
            </div>
          </form>
        </div>
      )}

      {/* EXCEL IMPORT MODAL */}
      {importModal && (
        <div
          className="overlay"
          onMouseDown={(e) => {
            if (e.target === e.currentTarget) setImportModal(false);
          }}
        >
          <section className="modal importmodal">
            <div className="modalhead">
              <div>
                <h2>استيراد العملاء من Excel</h2>
                <p>ارفع ملف Excel أو CSV، ثم راجع البيانات قبل اعتمادها.</p>
              </div>
              <button className="iconbtn" onClick={() => setImportModal(false)}>
                <X />
              </button>
            </div>

            <div className="importinstructions">
              <FileSpreadsheet />
              <div>
                <b>الأعمدة المطلوبة في الصف الأول</b>
                <span>{CUSTOMER_COLUMNS.join(' | ')}</span>
              </div>
            </div>

            <label className="uploadarea">
              <Upload />
              <b>{importFileName || 'اضغط لاختيار ملف Excel أو CSV'}</b>
              <small>الصيغ المدعومة: .xlsx، .xls، .csv</small>
              <input
                type="file"
                accept=".xlsx,.xls,.csv"
                onChange={(e) => {
                  const f = e.target.files?.[0];
                  if (f) parseImportFile(f);
                }}
              />
            </label>

            {importError && <div className="importerror">{importError}</div>}

            {importRows && (
              <>
                <div className="importsummary">
                  تمت قراءة <b>{importRows.length}</b> صفًا. عند التطابق يتم تحديث بيانات العميل باستخدام CIF، أو رقم التليفون إذا كان CIF غير موجود. لن تُحذف المتابعات.
                </div>
                <div className="importpreview">
                  <table>
                    <thead>
                      <tr>
                        {CUSTOMER_COLUMNS.map((h) => (
                          <th key={h}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {importRows.slice(0, 8).map((c) => (
                        <tr key={c.id}>
                          <td>{c.phone}</td>
                          <td>{c.name}</td>
                          <td>{c.installment ?? ''}</td>
                          <td>{c.daysLate ?? ''}</td>
                          <td>{c.dueDate}</td>
                          <td>{c.totalOverdue ?? ''}</td>
                          <td>{c.loanBalance ?? ''}</td>
                          <td>{c.currentAccountBalance ?? ''}</td>
                          <td>{c.currentAccount}</td>
                          <td>{c.cif}</td>
                          <td>{c.address}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                {importRows.length > 8 && (
                  <small className="previewhint">
                    معاينة أول 8 صفوف من أصل {importRows.length} صف.
                  </small>
                )}
              </>
            )}

            <div className="modalfoot">
              <button className="secondary" onClick={() => setImportModal(false)}>
                إلغاء
              </button>
              <button
                className="primary"
                disabled={!importRows?.length}
                onClick={confirmImport}
              >
                <Upload size={17} /> اعتماد الاستيراد
              </button>
            </div>
          </section>
        </div>
      )}

      {/* CONFIRMATION MODAL */}
      {deleteConfirm && (
        <div
          className="overlay"
          onMouseDown={(e) => {
            if (e.target === e.currentTarget) setDeleteConfirm(null);
          }}
        >
          <div className="modal smallmodal">
            <div className="modalhead">
              <div className="flex items-center gap-2">
                <div className="w-10 h-10 rounded-full bg-red-100 text-red-600 grid place-items-center shrink-0">
                  <AlertTriangle size={20} />
                </div>
                <div>
                  <h2 className="text-red-700">{deleteConfirm.title}</h2>
                  <p>{deleteConfirm.message}</p>
                </div>
              </div>
              <button
                type="button"
                className="iconbtn"
                onClick={() => setDeleteConfirm(null)}
              >
                <X size={18} />
              </button>
            </div>

            <div className="modalfoot">
              <button
                type="button"
                className="secondary"
                onClick={() => setDeleteConfirm(null)}
              >
                تراجع
              </button>
              <button
                className="dangerbtn"
                onClick={() => {
                  if (deleteConfirm.type === 'task' && deleteConfirm.id) {
                    deleteTask(deleteConfirm.id);
                  } else if (deleteConfirm.type === 'customer' && deleteConfirm.id) {
                    deleteCustomer(deleteConfirm.id);
                  } else if (deleteConfirm.type === 'clear_all') {
                    clearAllData();
                  }
                }}
              >
                {deleteConfirm.type === 'clear_all'
                  ? 'تأكيد مسح كافة البيانات نهائياً'
                  : 'تأكيد الإجراء'}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function Stat({
  icon,
  label,
  value
}: {
  icon: React.ReactNode;
  label: string;
  value: number;
}) {
  return (
    <div className="stat">
      <div className="staticon">{icon}</div>
      <div>
        <small>{label}</small>
        <b>{value}</b>
      </div>
    </div>
  );
}

function StatusBadge({ status }: { status: Status }) {
  return <span className={'status ' + statusClass(status)}>{status}</span>;
}

function TaskTable({
  tasks,
  customer,
  onEdit,
  onDelete,
  onStatus,
  onCustomer
}: {
  tasks: Task[];
  customer: (id: string) => Customer | undefined;
  onEdit: (t: Task) => void;
  onDelete?: (id: string) => void;
  onStatus: (id: string, s: Status) => void;
  onCustomer: (id: string) => void;
}) {
  if (!tasks.length)
    return (
      <div className="empty">
        <CalendarDays />
        <b>لا توجد مهام في هذا العرض</b>
        <span>يمكنك إضافة متابعة جديدة من الزر بالأعلى.</span>
      </div>
    );

  return (
    <div className="tablewrap">
      <table>
        <thead>
          <tr>
            <th>التاريخ والوقت</th>
            <th>العميل</th>
            <th>نوع المهمة</th>
            <th>الملاحظات</th>
            <th>الحالة</th>
            <th>إجراء</th>
          </tr>
        </thead>
        <tbody>
          {tasks.map((t) => {
            const cust = customer(t.customerId);
            return (
              <tr key={t.id}>
                <td>
                  <b>{t.date}</b>
                  <small className="flex items-center gap-1 font-semibold text-gray-500">
                    <Clock size={11} /> {t.time}
                  </small>
                </td>
                <td>
                  <button className="linkbutton" onClick={() => onCustomer(t.customerId)}>
                    {cust?.name || 'عميل'}
                  </button>
                  <small dir="ltr">{cust?.phone}</small>
                  {cust?.cif && <small className="text-gray-400">CIF: {cust.cif}</small>}
                </td>
                <td>
                  <span className="font-semibold">{t.kind}</span>
                </td>
                <td className="notecell">{t.notes || '—'}</td>
                <td>
                  <select
                    className={'statusselect ' + statusClass(t.status)}
                    value={t.status}
                    onChange={(e) => onStatus(t.id, e.target.value as Status)}
                  >
                    {['جديدة', 'قيد التنفيذ', 'تمت', 'مؤجلة'].map((s) => (
                      <option key={s}>{s}</option>
                    ))}
                  </select>
                </td>
                <td>
                  <div className="flex items-center gap-2">
                    <button className="textbutton" onClick={() => onEdit(t)}>
                      تعديل
                    </button>
                    {onDelete && (
                      <button
                        className="dangertextbtn"
                        onClick={() => onDelete(t.id)}
                        title="حذف المتابعة"
                      >
                        <Trash2 size={13} />
                      </button>
                    )}
                  </div>
                </td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}
