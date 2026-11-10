import React, { useState, useRef, useMemo, useEffect } from 'react';
import * as xlsx from 'xlsx';
import {
  UploadCloud, Download, Search, AlertCircle, Activity,
  ExternalLink, Mail, Phone, Loader2, Bookmark, BookmarkCheck,
  List, DatabaseZap, Trash2, Pencil, Check, X, ArrowRight,
  Building2, Sparkles, RefreshCw, ChevronDown
} from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Badge } from '@/components/ui/badge';
import { Progress } from '@/components/ui/progress';
import { Card, CardContent } from '@/components/ui/card';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select';
import { Tooltip, TooltipContent, TooltipTrigger } from '@/components/ui/tooltip';
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs';
import { Popover, PopoverContent, PopoverTrigger } from '@/components/ui/popover';
import { Checkbox } from '@/components/ui/checkbox';
import {
  Table, TableBody, TableCell, TableHead, TableHeader, TableRow,
} from '@/components/ui/table';
import { useToast } from '@/hooks/use-toast';

// ─── Types ──────────────────────────────────────────────────────────────────

type StudyContact = {
  name?: string;
  email?: string;
  phone?: string;
  place?: string;
};

type StudyData = {
  nctId: string;
  briefTitle?: string;
  overallStatus?: string;
  leadSponsor?: string;
  collaborators?: string[];
  therapeuticArea?: string;   // derived from conditions[0]
  indication?: string;         // all conditions joined
  interventions?: string[];
  phase?: string;
  enrollmentCount?: number;
  locationCount?: number;
  startDate?: string;
  origin?: string;             // not in API — user-editable
  contactName?: string;
  contactPhone?: string;
  contactEmail?: string;
  contacts?: StudyContact[];
  error?: boolean;
};

type SavedStudy = StudyData & {
  savedAt: string;
  origin: string;
};

const STORAGE_KEY = 'ct_saved_studies_v1';

// ─── Helpers ─────────────────────────────────────────────────────────────────

function loadSaved(): SavedStudy[] {
  try {
    return JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]');
  } catch {
    return [];
  }
}

function persistSaved(list: SavedStudy[]) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(list));
    return true;
  } catch {
    return false;
  }
}

function getStatusColor(status?: string) {
  if (!status) return 'bg-gray-100 text-gray-700 border-gray-200';
  const s = status.toUpperCase();
  if (s === 'RECRUITING') return 'bg-emerald-100 text-emerald-800 border-emerald-200';
  if (s === 'COMPLETED') return 'bg-slate-100 text-slate-700 border-slate-200';
  if (s === 'ACTIVE, NOT RECRUITING' || s === 'ACTIVE_NOT_RECRUITING') return 'bg-amber-100 text-amber-800 border-amber-200';
  if (s === 'TERMINATED' || s === 'WITHDRAWN' || s === 'SUSPENDED')
    return 'bg-red-100 text-red-800 border-red-200';
  if (s === 'NOT YET RECRUITING' || s === 'NOT_YET_RECRUITING' || s === 'ENROLLING_BY_INVITATION') return 'bg-blue-100 text-blue-800 border-blue-200';
  if (s === 'ERROR') return 'bg-destructive/10 text-destructive border-destructive/20';
  return 'bg-gray-100 text-gray-700 border-gray-200';
}

type SearchParams = {
  condition: string[];
  drug: string[];
  countries: string[];
  sponsors: string[];
  status: string;
  studyType: string;
  phase: string;
  sex: string;
  age: string;
  sponsorClass: string;
  interventionType: string;
  results: string;
};

type ApiContact = { name?: string; phone?: string; email?: string };
type ApiLocation = {
  facility?: string;
  city?: string;
  country?: string;
  status?: string;
  contacts?: ApiContact[];
};

const STUDY_FIELDS = [
  'NCTId', 'BriefTitle', 'OverallStatus', 'StartDate', 'LeadSponsorName', 'CollaboratorName',
  'Condition', 'InterventionName', 'Phase', 'EnrollmentCount', 'LocationFacility', 'LocationCity',
  'LocationCountry', 'LocationStatus', 'CentralContactName', 'CentralContactPhone', 'CentralContactEMail',
  'LocationContactName', 'LocationContactPhone', 'LocationContactEMail',
].join(',');

const ANY = 'ANY';

const SEARCH_STATUSES = [
  { value: ANY, label: 'Todos os status' },
  { value: 'RECRUITING', label: 'Recrutando' },
  { value: 'NOT_YET_RECRUITING', label: 'Ainda não recrutando' },
  { value: 'ENROLLING_BY_INVITATION', label: 'Por convite' },
  { value: 'ACTIVE_NOT_RECRUITING', label: 'Ativo, sem recrutar' },
  { value: 'COMPLETED', label: 'Concluído' },
  { value: 'SUSPENDED', label: 'Suspenso' },
  { value: 'TERMINATED', label: 'Encerrado' },
  { value: 'WITHDRAWN', label: 'Retirado' },
  { value: 'UNKNOWN', label: 'Desconhecido' },
  { value: 'AVAILABLE', label: 'Disponível' },
  { value: 'NO_LONGER_AVAILABLE', label: 'Não mais disponível' },
  { value: 'TEMPORARILY_NOT_AVAILABLE', label: 'Temporariamente indisponível' },
  { value: 'APPROVED_FOR_MARKETING', label: 'Aprovado para comercialização' },
  { value: 'WITHHELD', label: 'Retido' },
];

const SEARCH_STUDY_TYPES = [
  { value: ANY, label: 'Todos' },
  { value: 'INTERVENTIONAL', label: 'Intervencionista' },
  { value: 'OBSERVATIONAL', label: 'Observacional' },
  { value: 'EXPANDED_ACCESS', label: 'Acesso expandido' },
];

const SEARCH_PHASES = [
  { value: ANY, label: 'Qualquer' },
  { value: 'EARLY_PHASE1', label: 'Fase inicial 1' },
  { value: 'PHASE1', label: 'Fase 1' },
  { value: 'PHASE2', label: 'Fase 2' },
  { value: 'PHASE3', label: 'Fase 3' },
  { value: 'PHASE4', label: 'Fase 4' },
  { value: 'NA', label: 'Não se aplica' },
];

const SEARCH_SEXES = [
  { value: ANY, label: 'Qualquer' },
  { value: 'ALL', label: 'Todos os sexos' },
  { value: 'FEMALE', label: 'Feminino' },
  { value: 'MALE', label: 'Masculino' },
];

const SEARCH_AGES = [
  { value: ANY, label: 'Qualquer' },
  { value: 'CHILD', label: 'Criança' },
  { value: 'ADULT', label: 'Adulto' },
  { value: 'OLDER_ADULT', label: '65 anos ou mais' },
];

const SEARCH_SPONSOR_CLASSES = [
  { value: ANY, label: 'Qualquer' },
  { value: 'INDUSTRY', label: 'Indústria' },
  { value: 'NIH', label: 'NIH' },
  { value: 'FED', label: 'Governo federal dos EUA' },
  { value: 'OTHER_GOV', label: 'Outro governo' },
  { value: 'NETWORK', label: 'Rede' },
  { value: 'INDIV', label: 'Investigador' },
  { value: 'OTHER', label: 'Outro' },
  { value: 'UNKNOWN', label: 'Desconhecido' },
  { value: 'AMBIG', label: 'Ambíguo' },
];

const SEARCH_INTERVENTIONS = [
  { value: ANY, label: 'Todos' },
  { value: 'DRUG', label: 'Droga' },
  { value: 'BIOLOGICAL', label: 'Biológico' },
  { value: 'DEVICE', label: 'Dispositivo' },
  { value: 'PROCEDURE', label: 'Procedimento' },
  { value: 'RADIATION', label: 'Radiação' },
  { value: 'BEHAVIORAL', label: 'Comportamental' },
  { value: 'GENETIC', label: 'Genético' },
  { value: 'DIETARY_SUPPLEMENT', label: 'Suplemento' },
  { value: 'COMBINATION_PRODUCT', label: 'Combinação' },
  { value: 'DIAGNOSTIC_TEST', label: 'Teste diagnóstico' },
  { value: 'OTHER', label: 'Outro' },
];

const SEARCH_RESULTS = [
  { value: ANY, label: 'Qualquer' },
  { value: 'WITH', label: 'Com resultados' },
  { value: 'WITHOUT', label: 'Sem resultados' },
];

const COUNTRIES = [
  'Argentina', 'Bolivia', 'Brazil', 'Chile', 'Colombia', 'Costa Rica', 'Dominican Republic',
  'Ecuador', 'El Salvador', 'Guatemala', 'Honduras', 'Mexico', 'Panama', 'Paraguay', 'Peru',
  'Puerto Rico', 'Uruguay', 'Venezuela',
  'Australia', 'Austria', 'Belgium', 'Bulgaria', 'Canada', 'China', 'Czechia', 'Denmark',
  'Egypt', 'Finland', 'France', 'Germany', 'Greece', 'Hungary', 'India', 'Indonesia',
  'Ireland', 'Israel', 'Italy', 'Japan', 'Malaysia', 'Netherlands', 'New Zealand', 'Norway',
  'Philippines', 'Poland', 'Portugal', 'Romania', 'Saudi Arabia', 'Singapore', 'South Africa',
  'South Korea', 'Spain', 'Sweden', 'Switzerland', 'Taiwan', 'Thailand', 'Turkey (Türkiye)',
  'Ukraine', 'United Arab Emirates', 'United Kingdom', 'United States', 'Vietnam',
];

const SPONSORS = [
  'Novo Nordisk', 'Eli Lilly', 'Pfizer', 'Roche', 'Novartis', 'Sanofi', 'AstraZeneca',
  'Merck', 'AbbVie', 'Johnson & Johnson', 'Janssen', 'GSK', 'Bayer', 'Boehringer Ingelheim',
  'Amgen', 'Gilead', 'Bristol Myers Squibb', 'Takeda', 'Biogen', 'Regeneron', 'Moderna',
  'Vertex', 'Daiichi Sankyo', 'Astellas', 'Eisai', 'Teva', 'UCB', 'Ipsen', 'Ferring',
  'Chiesi', 'Servier', 'Organon', 'CSL', 'Lundbeck', 'Menarini',
];

const CONDITIONS = [
  'Type 2 Diabetes', 'Type 1 Diabetes', 'Obesity', 'Chronic Kidney Disease', 'Heart Failure',
  'Hypertension', 'Coronary Artery Disease', 'Atrial Fibrillation', 'Nonalcoholic Steatohepatitis',
  'Alzheimer Disease', 'Parkinson Disease', 'Migraine', 'Depression', 'Schizophrenia',
  'Breast Cancer', 'Lung Cancer', 'Prostate Cancer', 'Colorectal Cancer', 'Melanoma',
  'Multiple Myeloma', 'Lymphoma', 'Leukemia', 'Asthma', 'COPD', 'Atopic Dermatitis',
  'Psoriasis', 'Rheumatoid Arthritis', 'Ulcerative Colitis', 'Crohn Disease', 'Osteoporosis',
  'Osteoarthritis', 'HIV', 'COVID-19',
];

const DRUGS = [
  'Semaglutide', 'Tirzepatide', 'Dulaglutide', 'Liraglutide', 'Insulin', 'Metformin',
  'Empagliflozin', 'Dapagliflozin', 'Pembrolizumab', 'Nivolumab', 'Atezolizumab',
  'Durvalumab', 'Trastuzumab', 'Bevacizumab', 'Osimertinib', 'Adalimumab', 'Dupilumab',
  'Ustekinumab', 'Secukinumab', 'Risankizumab', 'Upadacitinib', 'Tocilizumab',
  'Apixaban', 'Rivaroxaban', 'Atorvastatin', 'Rosuvastatin', 'Lecanemab', 'Donanemab',
];

const PAGE_SIZES = [100, 200, 300, 400, 500, 600, 700, 800, 900, 1000];

function selectionLabel(selected: string[], emptyLabel: string) {
  if (selected.length === 0) return emptyLabel;
  if (selected.length === 1) return selected[0];
  return `${selected[0]} +${selected.length - 1}`;
}

function OptionPicker({
  options,
  selected,
  onChange,
  allowCustom = true,
  allowAll = false,
}: {
  options: string[];
  selected: string[];
  onChange: (values: string[]) => void;
  allowCustom?: boolean;
  allowAll?: boolean;
}) {
  const [draft, setDraft] = useState('');
  const [filter, setFilter] = useState('');
  const [extras, setExtras] = useState<string[]>([]);
  const choices = [
    ...options,
    ...extras.filter(name => !options.some(item => item.toLowerCase() === name.toLowerCase())),
  ];
  const visible = choices.filter(name => name.toLowerCase().includes(filter.trim().toLowerCase()));

  const add = () => {
    const name = draft.trim();
    if (!name) return;
    const existing = choices.find(item => item.toLowerCase() === name.toLowerCase());
    const value = existing ?? name;
    if (!existing) setExtras(current => [...current, name]);
    if (!selected.some(item => item.toLowerCase() === value.toLowerCase())) onChange([...selected, value]);
    setDraft('');
    setFilter('');
  };

  return (
    <Popover>
      <PopoverTrigger asChild>
        <Button type="button" variant="outline" className="h-10 w-full justify-between gap-2 px-3 font-normal">
          <span className={`truncate ${selected.length === 0 && !allowAll ? 'text-slate-400' : 'text-slate-900'}`}>
            {selectionLabel(selected, allowAll ? 'Todos' : 'Escolha um ou mais')}
          </span>
          <ChevronDown className="h-4 w-4 shrink-0 opacity-50" />
        </Button>
      </PopoverTrigger>
      <PopoverContent align="start" className="w-64 p-2">
        <Input
          value={filter}
          onChange={event => setFilter(event.target.value)}
          placeholder="Buscar na lista"
          className="mb-1.5 h-8 text-sm"
        />
        <div className="max-h-40 space-y-0.5 overflow-y-auto">
          {allowAll && (
            <label className="flex cursor-pointer items-center gap-2 rounded-md px-1.5 py-1 text-[13px] font-medium leading-4 hover:bg-[#f6f1ea]">
              <Checkbox
                checked={selected.length === 0}
                onCheckedChange={() => onChange([])}
              />
              <span>Todos</span>
            </label>
          )}
          {visible.map(name => (
            <label key={name} className="flex cursor-pointer items-center gap-2 rounded-md px-1.5 py-1 text-[13px] leading-4 hover:bg-[#f6f1ea]">
              <Checkbox
                checked={selected.includes(name)}
                onCheckedChange={checked => {
                  onChange(checked ? [...selected, name] : selected.filter(item => item !== name));
                }}
              />
              <span className="truncate">{name}</span>
            </label>
          ))}
          {visible.length === 0 && <p className="px-1.5 py-2 text-xs text-slate-400">Nada na lista</p>}
        </div>
        {allowCustom && (
          <div className="mt-2 flex gap-1.5">
            <Input
              value={draft}
              onChange={event => setDraft(event.target.value)}
              onKeyDown={event => {
                if (event.key === 'Enter') {
                  event.preventDefault();
                  add();
                }
              }}
              placeholder="Outro nome"
              className="h-8 text-sm"
            />
            <Button type="button" variant="outline" className="h-8 px-2 text-xs" onClick={add}>Incluir</Button>
          </div>
        )}
      </PopoverContent>
    </Popover>
  );
}

function locationQuery(countries: string[]) {
  return countries
    .map(country => (/[\s()]/.test(country) ? `"${country}"` : country))
    .join(' OR ');
}

const SEARCH_DROPDOWNS: { key: 'status' | 'studyType' | 'phase' | 'interventionType' | 'sex' | 'age' | 'sponsorClass' | 'results'; label: string; options: { value: string; label: string }[] }[] = [
  { key: 'status', label: 'Status', options: SEARCH_STATUSES },
  { key: 'studyType', label: 'Tipo de estudo', options: SEARCH_STUDY_TYPES },
  { key: 'phase', label: 'Fase', options: SEARCH_PHASES },
  { key: 'interventionType', label: 'Tipo de intervenção', options: SEARCH_INTERVENTIONS },
  { key: 'sex', label: 'Sexo', options: SEARCH_SEXES },
  { key: 'age', label: 'Idade', options: SEARCH_AGES },
  { key: 'sponsorClass', label: 'Classe do patrocinador', options: SEARCH_SPONSOR_CLASSES },
  { key: 'results', label: 'Resultados', options: SEARCH_RESULTS },
];

const SITE_STATUS: Record<string, string> = {
  RECRUITING: 'Recrutando',
  NOT_YET_RECRUITING: 'Ainda não recrutando',
  ACTIVE_NOT_RECRUITING: 'Ativo, sem recrutar',
  ENROLLING_BY_INVITATION: 'Por convite',
  COMPLETED: 'Concluído',
  SUSPENDED: 'Suspenso',
  TERMINATED: 'Encerrado',
  WITHDRAWN: 'Retirado',
};

function siteStatusLabel(status?: string) {
  if (!status) return '';
  return SITE_STATUS[status.toUpperCase().replace(/[\s,]+/g, '_')] || '';
}

function collectContacts(mod: { centralContacts?: ApiContact[]; locations?: ApiLocation[] }) {
  const out: StudyContact[] = [];
  const seen = new Set<string>();
  const push = (contact: StudyContact) => {
    const email = (contact.email || '').trim();
    const phone = (contact.phone || '').trim();
    if (!email && !phone) return;
    const key = `${email.toLowerCase()}|${phone}|${(contact.place || '').toLowerCase()}`;
    if (seen.has(key)) return;
    seen.add(key);
    out.push({ ...contact, name: (contact.name || '').trim(), email, phone });
  };

  for (const contact of mod.centralContacts || []) {
    push({
      name: contact.name,
      email: contact.email,
      phone: contact.phone,
      place: 'Contato central',
    });
  }

  for (const location of mod.locations || []) {
    const where = [location.facility, location.city, location.country].filter(Boolean).join(', ');
    const place = [where, siteStatusLabel(location.status)].filter(Boolean).join(' · ');
    for (const contact of location.contacts || []) {
      push({
        name: contact.name,
        email: contact.email,
        phone: contact.phone,
        place: place || 'Centro',
      });
    }
  }

  return out;
}

function focusQueries(focus?: string | string[]) {
  return (Array.isArray(focus) ? focus : focus ? [focus] : [])
    .map(item => item.trim().toLowerCase())
    .filter(Boolean);
}

function rankContacts(contacts: StudyContact[], focus?: string | string[]) {
  const queries = focusQueries(focus);
  const score = (contact: StudyContact) => {
    const place = (contact.place || '').toLowerCase();
    if (queries.some(query => place.includes(query))) return 0;
    if (place.startsWith('contato central')) return 1;
    return 2;
  };
  return [...contacts].sort((a, b) => score(a) - score(b) || Number(!a.email) - Number(!b.email));
}

function contactsOf(study: StudyData, focus?: string | string[]) {
  const source = study.contacts && study.contacts.length > 0
    ? study.contacts
    : (study.contactName || study.contactEmail || study.contactPhone)
      ? [{
          name: study.contactName,
          email: study.contactEmail,
          phone: study.contactPhone,
          place: 'Contato central',
        }]
      : [];
  const queries = focusQueries(focus);
  const scoped = queries.length === 0
    ? source
    : source.filter(contact => {
        const place = (contact.place || '').toLowerCase();
        if (place.startsWith('contato central')) return true;
        return queries.some(query => place.includes(query));
      });
  return rankContacts(scoped, focus);
}

function mapStudy(json: any, fallbackId: string): StudyData {
  const mod = json.protocolSection || {};
  const identification = mod.identificationModule || {};
  const status = mod.statusModule || {};
  const sponsor = mod.sponsorCollaboratorsModule || {};
  const conditions: string[] = mod.conditionsModule?.conditions || [];
  const arms = mod.armsInterventionsModule || {};
  const design = mod.designModule || {};
  const contactsMod = mod.contactsLocationsModule || {};
  const locations: ApiLocation[] = contactsMod.locations || [];
  const contacts = collectContacts(contactsMod);
  const primary = contacts.find(contact => contact.email) || contacts[0];
  const phases: string[] = design.phases || [];
  const interventionList = (arms.interventions || [])
    .map((item: { name?: string }) => item.name)
    .filter((name): name is string => !!name);
  const collaboratorList = (sponsor.collaborators || [])
    .map((item: { name?: string }) => item.name)
    .filter((name): name is string => !!name);

  return {
    nctId: identification.nctId || fallbackId,
    briefTitle: identification.briefTitle || '',
    overallStatus: status.overallStatus || '',
    startDate: status.startDateStruct?.date || '',
    leadSponsor: sponsor.leadSponsor?.name || '',
    collaborators: collaboratorList,
    therapeuticArea: conditions[0] || '',
    indication: conditions.join('; '),
    interventions: interventionList,
    phase: phases.join(', ').replace('PHASE', 'Fase'),
    enrollmentCount: design.enrollmentInfo?.count ?? undefined,
    locationCount: locations.length || undefined,
    contactName: primary?.name || '',
    contactPhone: primary?.phone || '',
    contactEmail: primary?.email || '',
    contacts,
    origin: '',
    error: false,
  };
}

async function fetchStudy(nctId: string): Promise<StudyData> {
  const res = await fetch(
    `https://clinicaltrials.gov/api/v2/studies/${encodeURIComponent(nctId)}?fields=${STUDY_FIELDS}`,
  );
  if (!res.ok) throw new Error('Not ok');
  return mapStudy(await res.json(), nctId);
}

async function queryStudies(params: SearchParams, pageToken?: string, pageSize = 100) {
  const url = new URL('https://clinicaltrials.gov/api/v2/studies');
  if (params.condition.length > 0) url.searchParams.set('query.cond', locationQuery(params.condition));
  if (params.drug.length > 0) url.searchParams.set('query.intr', locationQuery(params.drug));
  if (params.countries.length > 0) url.searchParams.set('query.locn', locationQuery(params.countries));
  if (params.sponsors.length > 0) url.searchParams.set('query.lead', locationQuery(params.sponsors));
  if (params.status && params.status !== ANY) url.searchParams.set('filter.overallStatus', params.status);
  const advanced = [
    params.studyType !== ANY ? `AREA[StudyType]${params.studyType}` : '',
    params.phase !== ANY ? `AREA[Phase]${params.phase}` : '',
    params.sex !== ANY ? `AREA[Sex]${params.sex}` : '',
    params.age !== ANY ? `AREA[StdAge]${params.age}` : '',
    params.sponsorClass !== ANY ? `AREA[LeadSponsorClass]${params.sponsorClass}` : '',
    params.interventionType !== ANY ? `AREA[InterventionType]${params.interventionType}` : '',
    params.results === 'WITH' ? 'AREA[HasResults]true' : '',
    params.results === 'WITHOUT' ? 'AREA[HasResults]false' : '',
  ].filter(Boolean);
  if (advanced.length > 0) url.searchParams.set('filter.advanced', advanced.join(' AND '));
  url.searchParams.set('fields', STUDY_FIELDS);
  url.searchParams.set('pageSize', String(pageSize));
  if (!pageToken) url.searchParams.set('countTotal', 'true');
  if (pageToken) url.searchParams.set('pageToken', pageToken);

  const res = await fetch(url);
  if (!res.ok) throw new Error('search failed');
  const json = await res.json();
  const studies = ((json.studies || []) as any[])
    .map(study => mapStudy(study, study?.protocolSection?.identificationModule?.nctId || ''))
    .filter(study => study.nctId);

  return {
    studies,
    totalCount: typeof json.totalCount === 'number' ? json.totalCount as number : undefined,
    nextPageToken: typeof json.nextPageToken === 'string' ? json.nextPageToken as string : undefined,
  };
}

// ─── Shared table component ───────────────────────────────────────────────────

function TruncatedCell({ value, maxW = 'max-w-[160px]' }: { value?: string; maxW?: string }) {
  if (!value) return <span className="text-slate-400">-</span>;
  return (
    <Tooltip>
      <TooltipTrigger asChild>
        <span className={`block truncate ${maxW} cursor-default`}>{value}</span>
      </TooltipTrigger>
      <TooltipContent side="bottom" className="max-w-sm leading-relaxed p-3 text-xs">
        {value}
      </TooltipContent>
    </Tooltip>
  );
}

function formatStudyDate(value?: string) {
  if (!value) return '';
  const match = value.match(/^(\d{4})(?:-(\d{2})(?:-(\d{2}))?)?/);
  if (!match) return value;
  const [, year, month, day] = match;
  if (day && month) return `${day}/${month}/${year}`;
  if (month) return `${month}/${year}`;
  return year;
}

function ContactCell({ contacts }: { contacts: StudyContact[] }) {
  if (contacts.length === 0) return <span className="text-xs text-slate-400">-</span>;

  return (
    <div className="max-h-32 w-[250px] space-y-2 overflow-y-auto rounded-lg border border-[#ece3d8] bg-[#fcfbf9] p-2">
      {contacts.map((contact, index) => (
        <div key={`${contact.email || ''}-${contact.phone || ''}-${index}`} className="space-y-0.5">
          {contact.name && (
            <div className="text-xs font-medium text-slate-900 dark:text-slate-100">{contact.name}</div>
          )}
          {contact.place && <div className="text-[11px] leading-4 text-slate-500">{contact.place}</div>}
          {contact.email && (
            <a href={`mailto:${contact.email}`} className="flex items-center gap-1 text-xs text-primary hover:underline">
              <Mail className="h-3 w-3 shrink-0 opacity-70" />
              <span className="break-all">{contact.email}</span>
            </a>
          )}
          {contact.phone && (
            <div className="flex items-center gap-1 text-xs text-slate-500">
              <Phone className="h-3 w-3 shrink-0 opacity-70" />
              {contact.phone}
            </div>
          )}
        </div>
      ))}
    </div>
  );
}

function StudyTable({
  studies,
  isFetching,
  actionSlot,
  contactFocus,
}: {
  studies: StudyData[];
  isFetching?: boolean;
  actionSlot: (study: StudyData) => React.ReactNode;
  contactFocus?: string | string[];
}) {
  const [searchQuery, setSearchQuery] = useState('');
  const [statusFilter, setStatusFilter] = useState('Todos');

  const uniqueStatuses = useMemo(() => {
    const s = new Set(studies.map(s => s.overallStatus).filter((x): x is string => !!x));
    return ['Todos', ...Array.from(s)];
  }, [studies]);

  const filtered = useMemo(() => studies.filter(s => {
    const q = searchQuery.toLowerCase();
    const matchSearch =
      s.nctId.toLowerCase().includes(q) ||
      (s.leadSponsor || '').toLowerCase().includes(q) ||
      (s.indication || '').toLowerCase().includes(q) ||
      (s.interventions || []).join(' ').toLowerCase().includes(q) ||
      (s.briefTitle || '').toLowerCase().includes(q) ||
      (s.contactName || '').toLowerCase().includes(q) ||
      (s.contactEmail || '').toLowerCase().includes(q) ||
      (s.contacts || []).some(contact =>
        (contact.name || '').toLowerCase().includes(q) ||
        (contact.email || '').toLowerCase().includes(q) ||
        (contact.phone || '').toLowerCase().includes(q) ||
        (contact.place || '').toLowerCase().includes(q),
      );
    const matchStatus = statusFilter === 'Todos' || s.overallStatus === statusFilter;
    return matchSearch && matchStatus;
  }), [studies, searchQuery, statusFilter]);

  return (
    <div className="space-y-3">
      <div className="flex flex-col sm:flex-row gap-3 items-center bg-white dark:bg-slate-900 p-3 rounded-lg border shadow-sm">
        <div className="relative w-full sm:w-80">
          <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-slate-400" />
          <Input
            placeholder="Buscar estudos, patrocinadores, indicações..."
            value={searchQuery}
            onChange={e => setSearchQuery(e.target.value)}
            className="pl-9 bg-slate-50 dark:bg-slate-950 border-slate-200 dark:border-slate-800"
          />
        </div>
        <div className="flex items-center gap-2 w-full sm:w-auto">
          <span className="text-sm font-medium text-slate-500 whitespace-nowrap">Status:</span>
          <Select value={statusFilter} onValueChange={setStatusFilter}>
            <SelectTrigger className="w-full sm:w-[180px] bg-slate-50 dark:bg-slate-950 border-slate-200 dark:border-slate-800">
              <SelectValue placeholder="Todos" />
            </SelectTrigger>
            <SelectContent>
              {uniqueStatuses.map(s => (
                <SelectItem key={s} value={s}>{s === 'Todos' ? 'Todos' : s || 'Desconhecido'}</SelectItem>
              ))}
            </SelectContent>
          </Select>
        </div>
      </div>

      <div className="bg-white dark:bg-slate-900 rounded-lg border shadow-sm overflow-x-auto">
        <Table>
          <TableHeader className="bg-slate-50/70 dark:bg-slate-950/70">
            <TableRow className="hover:bg-transparent">
              <TableHead className="font-semibold text-slate-800 dark:text-slate-300 whitespace-nowrap">Área Terapêutica</TableHead>
              <TableHead className="font-semibold text-slate-800 dark:text-slate-300 whitespace-nowrap">Patrocinador</TableHead>
              <TableHead className="font-semibold text-slate-800 dark:text-slate-300 whitespace-nowrap">Origem (indicação)</TableHead>
              <TableHead className="font-semibold text-slate-800 dark:text-slate-300 whitespace-nowrap">Indicação</TableHead>
              <TableHead className="font-semibold text-slate-800 dark:text-slate-300 whitespace-nowrap">Intervenção</TableHead>
              <TableHead className="font-semibold text-slate-800 dark:text-slate-300 whitespace-nowrap">Fase</TableHead>
              <TableHead className="font-semibold text-slate-800 dark:text-slate-300 whitespace-nowrap">Status</TableHead>
              <TableHead className="font-semibold text-slate-800 dark:text-slate-300 whitespace-nowrap text-right">Nº de centros</TableHead>
              <TableHead className="font-semibold text-slate-800 dark:text-slate-300 whitespace-nowrap text-right">Nº de participantes</TableHead>
              <TableHead className="font-semibold text-slate-800 dark:text-slate-300 whitespace-nowrap">Início</TableHead>
              <TableHead className="font-semibold text-slate-800 dark:text-slate-300 whitespace-nowrap">NCT ID</TableHead>
              <TableHead className="font-semibold text-slate-800 dark:text-slate-300 whitespace-nowrap">CRO / Colaborador</TableHead>
              <TableHead className="font-semibold text-slate-800 dark:text-slate-300 whitespace-nowrap">Contato</TableHead>
              <TableHead className="w-[80px]" />
            </TableRow>
          </TableHeader>
          <TableBody>
            {filtered.length === 0 ? (
              <TableRow>
                <TableCell colSpan={14} className="h-40 text-center text-slate-500">
                  {isFetching ? 'Carregando primeiro lote...' : 'Nenhum resultado encontrado.'}
                </TableCell>
              </TableRow>
            ) : (
              filtered.map(study => (
                <TableRow
                  key={study.nctId}
                  className={study.error ? 'bg-red-50/30 dark:bg-red-950/10' : 'hover:bg-slate-50/50 dark:hover:bg-slate-800/30'}
                >
                  {/* Therapeutic Area */}
                  <TableCell className="text-sm"><TruncatedCell value={study.therapeuticArea} maxW="max-w-[140px]" /></TableCell>

                  {/* Sponsor */}
                  <TableCell className="text-sm"><TruncatedCell value={study.leadSponsor} maxW="max-w-[140px]" /></TableCell>

                  {/* Origin (indic.) — not from API */}
                  <TableCell className="text-sm text-slate-400 italic">-</TableCell>

                  {/* Indication */}
                  <TableCell className="text-sm"><TruncatedCell value={study.indication} maxW="max-w-[160px]" /></TableCell>

                  {/* Intervention */}
                  <TableCell className="text-sm">
                    <TruncatedCell value={(study.interventions || []).join('; ')} maxW="max-w-[140px]" />
                  </TableCell>

                  {/* Phase */}
                  <TableCell className="text-sm whitespace-nowrap">
                    {study.phase || <span className="text-slate-400">-</span>}
                  </TableCell>

                  {/* Status */}
                  <TableCell>
                    {study.error ? (
                      <span className="text-destructive text-xs flex items-center gap-1">
                        <AlertCircle className="h-3 w-3" /> Erro
                      </span>
                    ) : (
                      <Badge variant="outline" className={`text-xs font-medium border shadow-none whitespace-nowrap ${getStatusColor(study.overallStatus)}`}>
                        {study.overallStatus || 'Desconhecido'}
                      </Badge>
                    )}
                  </TableCell>

                  {/* No. of sites */}
                  <TableCell className="text-sm text-right tabular-nums">
                    {study.locationCount != null ? study.locationCount : <span className="text-slate-400">-</span>}
                  </TableCell>

                  {/* No. of participants */}
                  <TableCell className="text-sm text-right tabular-nums">
                    {study.enrollmentCount != null ? study.enrollmentCount.toLocaleString() : <span className="text-slate-400">-</span>}
                  </TableCell>

                  {/* Start */}
                  <TableCell className="text-sm whitespace-nowrap tabular-nums">
                    {formatStudyDate(study.startDate) || <span className="text-slate-400">-</span>}
                  </TableCell>

                  {/* NCT ID */}
                  <TableCell className="font-mono text-xs">
                    {study.error ? (
                      <span className="text-slate-500">{study.nctId}</span>
                    ) : (
                      <a
                        href={`https://clinicaltrials.gov/study/${study.nctId}`}
                        target="_blank"
                        rel="noopener noreferrer"
                        className="text-primary hover:underline underline-offset-4 inline-flex items-center gap-1"
                      >
                        {study.nctId}
                        <ExternalLink className="h-3 w-3 opacity-50" />
                      </a>
                    )}
                  </TableCell>

                  {/* CRO / Collaborator */}
                  <TableCell className="text-sm">
                    <TruncatedCell value={(study.collaborators || []).join('; ')} maxW="max-w-[140px]" />
                  </TableCell>

                  <TableCell>
                    {study.error ? <span className="text-xs text-slate-400">-</span> : <ContactCell contacts={contactsOf(study, contactFocus)} />}
                  </TableCell>

                  {/* Action */}
                  <TableCell>{actionSlot(study)}</TableCell>
                </TableRow>
              ))
            )}
          </TableBody>
        </Table>
      </div>

      {!isFetching && filtered.length > 0 && (
        <p className="text-center text-xs text-slate-400 py-2">
          {filtered.length} {filtered.length === 1 ? 'estudo' : 'estudos'}
        </p>
      )}
    </div>
  );
}

// ─── Saved list panel ─────────────────────────────────────────────────────────

function SavedList({
  saved,
  onRemove,
  onExport,
  onUpdateOrigin,
  onRefresh,
  isRefreshing,
}: {
  saved: SavedStudy[];
  onRemove: (nctId: string) => void;
  onExport: () => void;
  onUpdateOrigin: (nctId: string, origin: string) => void;
  onRefresh: () => void;
  isRefreshing: boolean;
}) {
  const [editingId, setEditingId] = useState<string | null>(null);
  const [editValue, setEditValue] = useState('');

  // We need to pass edits up — use a ref trick by passing a callback
  // This component is read-only for origin; inline edit handled via prop
  return (
    <div className="space-y-3">
      <div className="flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
        <p className="text-xs text-slate-500">Atualizar busca de novo os contatos e o status, sem reenviar a planilha.</p>
        <div className="flex justify-end gap-2">
          <Button size="sm" variant="outline" onClick={onRefresh} disabled={saved.length === 0 || isRefreshing} className="gap-2">
            {isRefreshing ? <Loader2 className="h-4 w-4 animate-spin" /> : <RefreshCw className="h-4 w-4" />}
            Atualizar contatos
          </Button>
          <Button size="sm" variant="outline" onClick={onExport} disabled={saved.length === 0} className="gap-2">
            <Download className="h-4 w-4" /> Exportar CSV
          </Button>
        </div>
      </div>

      {saved.length === 0 ? (
        <Card className="border-dashed border-2 border-slate-200 dark:border-slate-800 shadow-none">
          <CardContent className="flex flex-col items-center justify-center py-16 text-center">
            <Bookmark className="h-10 w-10 text-slate-300 mb-4" />
            <p className="text-slate-500 font-medium">Nenhum estudo salvo ainda.</p>
            <p className="text-slate-400 text-sm mt-1">Clique em "Salvar" em qualquer estudo da aba Resultados.</p>
          </CardContent>
        </Card>
      ) : (
        <div className="bg-white dark:bg-slate-900 rounded-lg border shadow-sm overflow-x-auto">
          <Table>
            <TableHeader className="bg-slate-50/70 dark:bg-slate-950/70">
              <TableRow className="hover:bg-transparent">
                <TableHead className="font-semibold text-slate-800 whitespace-nowrap">Área Terapêutica</TableHead>
                <TableHead className="font-semibold text-slate-800 whitespace-nowrap">Patrocinador</TableHead>
                <TableHead className="font-semibold text-slate-800 whitespace-nowrap">Origem (indicação)</TableHead>
                <TableHead className="font-semibold text-slate-800 whitespace-nowrap">Indicação</TableHead>
                <TableHead className="font-semibold text-slate-800 whitespace-nowrap">Intervenção</TableHead>
                <TableHead className="font-semibold text-slate-800 whitespace-nowrap">Fase</TableHead>
                <TableHead className="font-semibold text-slate-800 whitespace-nowrap">Status</TableHead>
                <TableHead className="font-semibold text-slate-800 whitespace-nowrap text-right">Nº de centros</TableHead>
                <TableHead className="font-semibold text-slate-800 whitespace-nowrap text-right">Nº de participantes</TableHead>
                <TableHead className="font-semibold text-slate-800 whitespace-nowrap">Início</TableHead>
                <TableHead className="font-semibold text-slate-800 whitespace-nowrap">NCT ID</TableHead>
                <TableHead className="font-semibold text-slate-800 whitespace-nowrap">CRO / Colaborador</TableHead>
                <TableHead className="font-semibold text-slate-800 whitespace-nowrap">Contato</TableHead>
                <TableHead className="w-[80px]" />
              </TableRow>
            </TableHeader>
            <TableBody>
              {saved.map(study => (
                <TableRow key={study.nctId} className="hover:bg-slate-50/50">
                  <TableCell className="text-sm"><TruncatedCell value={study.therapeuticArea} maxW="max-w-[120px]" /></TableCell>
                  <TableCell className="text-sm"><TruncatedCell value={study.leadSponsor} maxW="max-w-[120px]" /></TableCell>

                  {/* Origin — editable */}
                  <TableCell className="text-sm">
                    {editingId === study.nctId ? (
                      <div className="flex items-center gap-1">
                        <Input
                          value={editValue}
                          onChange={e => setEditValue(e.target.value)}
                          className="h-7 w-28 border-slate-200 bg-white/90 px-2 text-xs"
                          autoFocus
                        />
                        <Button size="icon" variant="ghost" className="h-6 w-6"
                          onClick={() => {
                            onUpdateOrigin(study.nctId, editValue.trim());
                            setEditingId(null);
                          }}>
                          <Check className="h-3 w-3 text-emerald-600" />
                        </Button>
                        <Button size="icon" variant="ghost" className="h-6 w-6"
                          onClick={() => setEditingId(null)}>
                          <X className="h-3 w-3 text-red-500" />
                        </Button>
                      </div>
                    ) : (
                      <span
                        className="flex items-center gap-1 cursor-pointer group"
                        onClick={() => { setEditingId(study.nctId); setEditValue(study.origin || ''); }}
                      >
                        <span className={study.origin ? 'text-slate-700' : 'text-slate-400 italic'}>
                          {study.origin || 'Editar...'}
                        </span>
                        <Pencil className="h-3 w-3 text-slate-300 group-hover:text-slate-500 transition-colors" />
                      </span>
                    )}
                  </TableCell>

                  <TableCell className="text-sm"><TruncatedCell value={study.indication} maxW="max-w-[140px]" /></TableCell>
                  <TableCell className="text-sm"><TruncatedCell value={(study.interventions || []).join('; ')} maxW="max-w-[120px]" /></TableCell>
                  <TableCell className="text-sm whitespace-nowrap">{study.phase || <span className="text-slate-400">-</span>}</TableCell>
                  <TableCell>
                    <Badge variant="outline" className={`text-xs border shadow-none whitespace-nowrap ${getStatusColor(study.overallStatus)}`}>
                      {study.overallStatus || '-'}
                    </Badge>
                  </TableCell>
                  <TableCell className="text-sm text-right tabular-nums">
                    {study.locationCount != null ? study.locationCount : <span className="text-slate-400">-</span>}
                  </TableCell>
                  <TableCell className="text-sm text-right tabular-nums">
                    {study.enrollmentCount != null ? study.enrollmentCount.toLocaleString() : <span className="text-slate-400">-</span>}
                  </TableCell>
                  <TableCell className="text-sm whitespace-nowrap">{formatStudyDate(study.startDate) || <span className="text-slate-400">-</span>}</TableCell>
                  <TableCell className="font-mono text-xs">
                    <a
                      href={`https://clinicaltrials.gov/study/${study.nctId}`}
                      target="_blank" rel="noopener noreferrer"
                      className="text-primary hover:underline underline-offset-4 inline-flex items-center gap-1"
                    >
                      {study.nctId}
                      <ExternalLink className="h-3 w-3 opacity-50" />
                    </a>
                  </TableCell>
                  <TableCell className="text-sm"><TruncatedCell value={(study.collaborators || []).join('; ')} maxW="max-w-[120px]" /></TableCell>
                  <TableCell>
                    <ContactCell contacts={contactsOf(study)} />
                  </TableCell>
                  <TableCell>
                    <Button
                      size="icon"
                      variant="ghost"
                      className="h-7 w-7 text-slate-400 hover:text-red-500"
                      onClick={() => onRemove(study.nctId)}
                    >
                      <Trash2 className="h-4 w-4" />
                    </Button>
                  </TableCell>
                </TableRow>
              ))}
            </TableBody>
          </Table>
        </div>
      )}
    </div>
  );
}

// ─── Main Dashboard ────────────────────────────────────────────────────────────

export default function Dashboard() {
  const { toast } = useToast();
  const fileInputRef = useRef<HTMLInputElement>(null);

  const [studies, setStudies] = useState<StudyData[]>([]);
  const [isUploading, setIsUploading] = useState(false);
  const [isFetching, setIsFetching] = useState(false);
  const [progress, setProgress] = useState({ current: 0, total: 0 });
  const [hasData, setHasData] = useState(false);
  const [activeTab, setActiveTab] = useState('results');
  const [savedStudies, setSavedStudies] = useState<SavedStudy[]>(loadSaved);
  const [isRefreshing, setIsRefreshing] = useState(false);
  const [isSearching, setIsSearching] = useState(false);
  const [contactFocus, setContactFocus] = useState<string[]>([]);
  const [searchTotal, setSearchTotal] = useState<number | null>(null);
  const [nextPageToken, setNextPageToken] = useState<string | undefined>();
  const [pageSize, setPageSize] = useState(100);
  const [searchForm, setSearchForm] = useState<SearchParams>({
    condition: [],
    drug: [],
    countries: [],
    sponsors: [],
    status: 'RECRUITING',
    studyType: ANY,
    phase: ANY,
    sex: ANY,
    age: ANY,
    sponsorClass: ANY,
    interventionType: ANY,
    results: ANY,
  });
  const lastQuery = useRef<SearchParams>(searchForm);
  const storageWarned = useRef(false);

  useEffect(() => {
    const savedOk = persistSaved(savedStudies);
    if (!savedOk && !storageWarned.current) {
      storageWarned.current = true;
      toast({
        title: 'A lista não coube no navegador',
        description: 'Os contatos seguem nesta tela, mas o navegador recusou gravar a lista completa.',
        variant: 'destructive',
      });
    }
  }, [savedStudies, toast]);

  const savedIds = useMemo(() => new Set(savedStudies.map(s => s.nctId)), [savedStudies]);
  const totalStudies = studies.length + savedStudies.length;
  const sponsorCount = useMemo(() => {
    const sponsors = new Set(
      [...studies, ...savedStudies]
        .map(study => study.leadSponsor?.trim())
        .filter((value): value is string => !!value),
    );
    return sponsors.size;
  }, [studies, savedStudies]);
  const recruitingCount = useMemo(
    () => [...studies, ...savedStudies].filter(study => study.overallStatus?.toUpperCase() === 'RECRUITING').length,
    [studies, savedStudies],
  );
  const highlightedStudies = useMemo(
    () => [...savedStudies, ...studies].filter(study => contactsOf(study).length > 0).length,
    [savedStudies, studies],
  );

  // ── File upload ──────────────────────────────────────────────────────────────
  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setIsUploading(true);

    try {
      const buffer = await file.arrayBuffer();
      const workbook = xlsx.read(buffer);
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const data = xlsx.utils.sheet_to_json<any[]>(sheet, { header: 1 });

      const nctIds = new Set<string>();
      for (const row of data) {
        if (Array.isArray(row)) {
          for (const cell of row) {
            if (typeof cell === 'string') {
              const match = cell.match(/NCT\d+/i);
              if (match) nctIds.add(match[0].toUpperCase());
            }
          }
        }
      }

      const allIds = Array.from(nctIds);
      if (allIds.length === 0) {
        toast({ title: 'Nenhum ID NCT encontrado no arquivo', variant: 'destructive' });
        return;
      }

      // Filter out IDs already saved
      const newIds = allIds.filter(id => !savedIds.has(id));
      const alreadySaved = allIds.length - newIds.length;

      if (alreadySaved > 0) {
        toast({
          title: `${alreadySaved} estudo(s) já estão na sua lista`,
          description: `Buscando apenas os ${newIds.length} novos.`,
        });
      }

      if (newIds.length === 0) {
        toast({ title: 'Todos os estudos já estão na sua lista salva.' });
        setActiveTab('saved');
        return;
      }

      setHasData(true);
      setActiveTab('results');
      setContactFocus([]);
      setNextPageToken(undefined);
      setSearchTotal(null);
      startFetching(newIds);

    } catch (err) {
      console.error(err);
      toast({ title: 'Erro ao processar o arquivo Excel', variant: 'destructive' });
    } finally {
      setIsUploading(false);
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  // ── Batch fetch ──────────────────────────────────────────────────────────────
  const startFetching = async (nctIds: string[]) => {
    setIsFetching(true);
    setProgress({ current: 0, total: nctIds.length });
    setStudies([]);

    const BATCH_SIZE = 5;
    for (let i = 0; i < nctIds.length; i += BATCH_SIZE) {
      const batch = nctIds.slice(i, i + BATCH_SIZE);
      const results = await Promise.all(
        batch.map(id =>
          fetchStudy(id).catch(() => ({
            nctId: id,
            error: true,
            briefTitle: 'Erro ao carregar',
            overallStatus: 'Erro',
          } as StudyData))
        )
      );
      setStudies(prev => [...prev, ...results]);
      setProgress(prev => ({
        ...prev,
        current: Math.min(prev.current + BATCH_SIZE, prev.total),
      }));
    }
    setIsFetching(false);
  };

  const runSearch = async () => {
    const hasText = searchForm.condition.length > 0 || searchForm.drug.length > 0 || searchForm.countries.length > 0 || searchForm.sponsors.length > 0;
    const hasDropdown = SEARCH_DROPDOWNS.some(filter => searchForm[filter.key] !== ANY);
    if (!hasText && !hasDropdown) {
      toast({ title: 'Informe um texto ou escolha um filtro', variant: 'destructive' });
      return;
    }
    setIsSearching(true);
    try {
      const page = await queryStudies(searchForm, undefined, pageSize);
      lastQuery.current = searchForm;
      setStudies(page.studies);
      setHasData(true);
      setActiveTab('results');
      setContactFocus(searchForm.countries);
      setSearchTotal(page.totalCount ?? page.studies.length);
      setNextPageToken(page.nextPageToken);
      if (page.studies.length === 0) toast({ title: 'Nenhum estudo encontrado para essa busca' });
    } catch (err) {
      console.error(err);
      toast({ title: 'Não foi possível consultar ClinicalTrials.gov', variant: 'destructive' });
    } finally {
      setIsSearching(false);
    }
  };

  const loadMore = async () => {
    if (!nextPageToken || isSearching) return;
    setIsSearching(true);
    try {
      const page = await queryStudies(lastQuery.current, nextPageToken, pageSize);
      setStudies(prev => {
        const seen = new Set(prev.map(study => study.nctId));
        return [...prev, ...page.studies.filter(study => !seen.has(study.nctId))];
      });
      setNextPageToken(page.nextPageToken);
    } catch (err) {
      console.error(err);
      toast({ title: 'Não foi possível carregar mais estudos', variant: 'destructive' });
    } finally {
      setIsSearching(false);
    }
  };

  const refreshSaved = async () => {
    if (savedStudies.length === 0 || isRefreshing) return;
    setIsRefreshing(true);
    setProgress({ current: 0, total: savedStudies.length });
    const snapshot = savedStudies;
    const BATCH_SIZE = 5;
    for (let i = 0; i < snapshot.length; i += BATCH_SIZE) {
      const batch = snapshot.slice(i, i + BATCH_SIZE);
      const results = await Promise.all(batch.map(async saved => {
        try {
          const fresh = await fetchStudy(saved.nctId);
          return { ...fresh, origin: saved.origin, savedAt: saved.savedAt };
        } catch {
          return saved;
        }
      }));
      setSavedStudies(prev => {
        const byId = new Map(results.map(study => [study.nctId, study]));
        return prev.map(study => byId.get(study.nctId) ?? study);
      });
      setProgress({ current: Math.min(i + BATCH_SIZE, snapshot.length), total: snapshot.length });
    }
    setIsRefreshing(false);
    toast({ title: 'Lista atualizada com contatos ao vivo.' });
  };

  // ── Save / remove ────────────────────────────────────────────────────────────
  const saveStudy = (study: StudyData) => {
    if (savedIds.has(study.nctId)) return;
    const entry: SavedStudy = { ...study, origin: '', savedAt: new Date().toISOString() };
    setSavedStudies(prev => [...prev, entry]);
    toast({ title: `${study.nctId} salvo na sua lista.` });
  };

  const removeStudy = (nctId: string) => {
    setSavedStudies(prev => prev.filter(s => s.nctId !== nctId));
    toast({ title: `${nctId} removido da lista.` });
  };

  const updateSavedOrigin = (nctId: string, origin: string) => {
    setSavedStudies(prev =>
      prev.map(study => (study.nctId === nctId ? { ...study, origin } : study)),
    );
    toast({ title: `${nctId} atualizado`, description: 'Campo de origem salvo com sucesso.' });
  };

  // ── CSV export ───────────────────────────────────────────────────────────────
  const escapeField = (v?: string | number | null) => {
    const s = v == null ? '' : String(v);
    return `"${s.replace(/"/g, '""')}"`;
  };

  const exportCSV = (list: StudyData[], focus?: string[]) => {
    const headers = [
      'NCT ID', 'Área Terapêutica', 'Patrocinador', 'Origem (indicação)', 'Indicação',
      'Intervenção', 'Fase', 'Status', 'Nº de centros', 'Nº de participantes',
      'Início', 'CRO / Colaborador', 'Nome do contato', 'E-mail do contato', 'Telefone do contato',
    ];
    const rows = list.map(d => [
      escapeField(d.nctId),
      escapeField(d.therapeuticArea),
      escapeField(d.leadSponsor),
      escapeField((d as SavedStudy).origin),
      escapeField(d.indication),
      escapeField((d.interventions || []).join('; ')),
      escapeField(d.phase),
      escapeField(d.overallStatus),
      escapeField(d.locationCount),
      escapeField(d.enrollmentCount),
      escapeField(formatStudyDate(d.startDate)),
      escapeField((d.collaborators || []).join('; ')),
      escapeField(contactsOf(d, focus).map(contact => [contact.name, contact.place].filter(Boolean).join(' — ')).join(' | ')),
      escapeField(contactsOf(d, focus).map(contact => contact.email || '').join(' | ')),
      escapeField(contactsOf(d, focus).map(contact => contact.phone || '').join(' | ')),
    ]);
    const csv = [headers.map(escapeField).join(','), ...rows.map(r => r.join(','))].join('\r\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `estudos_clinicos_${new Date().toISOString().split('T')[0]}.csv`;
    a.style.visibility = 'hidden';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  };

  // ── Render ───────────────────────────────────────────────────────────────────
  return (
    <div className="avanzare-stage relative min-h-screen overflow-hidden px-4 py-4 font-sans md:px-8 md:py-8">
      <div className="avanzare-orb -left-20 top-8 h-64 w-64 bg-blue-400/35" />
      <div className="avanzare-orb right-[-40px] top-24 h-72 w-72 bg-blue-800/25 [animation-delay:-4s]" />
      <div className="avanzare-orb bottom-0 left-1/3 h-56 w-56 bg-sky-300/40 [animation-delay:-8s]" />
      <div className="relative mx-auto max-w-[1600px] space-y-6">
        <Input
          type="file"
          accept=".xlsx,.xls"
          className="hidden"
          ref={fileInputRef}
          onChange={handleFileUpload}
        />

        <section className="avanzare-panel avanzare-rise rounded-[28px]">
          <div className="relative overflow-hidden border-b border-blue-100 bg-gradient-to-r from-[#0e3a6b] to-[#1d6fbf] px-6 py-6 text-white md:px-8 md:py-8">
            <video
              className="avanzare-hero-video"
              autoPlay
              muted
              loop
              playsInline
              src="/avanzare-hero.mp4"
            />
            <div className="avanzare-hero-wash pointer-events-none absolute inset-0" />
            <div className="relative flex flex-col gap-6 lg:flex-row lg:items-end lg:justify-between">
              <div className="max-w-4xl space-y-5">
                <div className="inline-flex rounded-[26px] border border-[#efe4d8] bg-white/88 px-5 py-4 shadow-[0_14px_30px_rgba(48,26,70,0.05)] backdrop-blur">
                  <img
                    src="/avanzare-logo.jpg"
                    alt="Avanzare Clinical Research Solutions"
                    className="h-auto w-[220px] md:w-[320px]"
                  />
                </div>
                <div className="avanzare-hero space-y-3">
                  <div className="text-[11px] font-semibold uppercase tracking-[0.26em] text-[#8d5f37]">
                    Avanzare Clinical Research Solutions
                  </div>
                  <h1 className="max-w-3xl text-3xl font-semibold tracking-[-0.045em] text-[#34204f] md:text-5xl">
                    Espaço de trabalho para prospecção de estudos clínicos
                  </h1>
                  <p className="max-w-2xl text-sm leading-7 text-slate-600 md:text-base">
                    Busque ao vivo por condição, droga, local ou patrocinador, ou envie uma planilha com NCT. As colunas são as mesmas; os contatos saem atualizados da API.
                  </p>
                </div>
              </div>

              <div className="flex items-end lg:justify-end">
                <div className="rounded-2xl border border-[#ecdccc] bg-white/88 px-4 py-3 text-left shadow-[0_10px_24px_rgba(48,26,70,0.04)] backdrop-blur lg:text-right">
                  <div className="text-[11px] font-semibold uppercase tracking-[0.22em] text-[#8d5f37]">
                    By Lisa Palla Tavolaro & Santiago J. Isbert Perlender
                  </div>
                </div>
              </div>
            </div>
          </div>

          <div className="grid gap-0 lg:grid-cols-[1.25fr_0.75fr]">
            <div className="space-y-6 px-6 py-6 md:px-8 md:py-8">
              <div className="grid gap-4 md:grid-cols-2 xl:grid-cols-4">
                {[
                  {
                    label: 'Estudos visíveis',
                    value: totalStudies,
                    note: 'Resultados e shortlist combinados',
                    icon: Activity,
                    card: 'border-blue-200 bg-white',
                    chip: 'bg-blue-100 text-[#0e3a6b]',
                    number: 'text-[#0e3a6b]',
                  },
                  {
                    label: 'Em recrutamento',
                    value: recruitingCount,
                    note: 'Registros prontos para priorização',
                    icon: Sparkles,
                    card: 'border-sky-200 bg-sky-50',
                    chip: 'bg-sky-100 text-sky-800',
                    number: 'text-sky-900',
                  },
                  {
                    label: 'Patrocinadores',
                    value: sponsorCount,
                    note: 'Organizações únicas no espaço',
                    icon: Building2,
                    card: 'border-blue-300 bg-blue-50',
                    chip: 'bg-blue-200 text-[#0e3a6b]',
                    number: 'text-[#0e3a6b]',
                  },
                  {
                    label: 'Com contatos',
                    value: highlightedStudies,
                    note: 'Entradas com e-mail ou telefone',
                    icon: Mail,
                    card: 'border-[#0e3a6b]/20 bg-white',
                    chip: 'bg-[#0e3a6b]/10 text-[#0e3a6b]',
                    number: 'text-[#0e3a6b]',
                  },
                ].map(metric => (
                  <Card key={metric.label} className={`rounded-2xl shadow-none transition duration-200 hover:-translate-y-0.5 hover:shadow-md ${metric.card}`}>
                    <CardContent className="p-5">
                      <div className="mb-4 flex items-center justify-between">
                        <div className="text-[11px] font-semibold uppercase tracking-[0.22em] text-slate-500">
                          {metric.label}
                        </div>
                        <div className={`rounded-full p-2 ${metric.chip}`}>
                          <metric.icon className="h-4 w-4" />
                        </div>
                      </div>
                      <div className={`text-3xl font-semibold tracking-[-0.04em] ${metric.number}`}>
                        {metric.value}
                      </div>
                      <p className="mt-2 text-sm leading-6 text-slate-500">{metric.note}</p>
                    </CardContent>
                  </Card>
                ))}
              </div>
            </div>

            <div className="border-t border-blue-100 bg-blue-50/70 px-6 py-6 lg:border-l lg:border-t-0 md:px-8 md:py-8">
              <div className="rounded-[24px] border border-blue-100 bg-white p-6">
                <div className="mb-6 flex items-start justify-between gap-4">
                  <div>
                    <div className="text-[11px] font-semibold uppercase tracking-[0.22em] text-[#8d5f37]">
                      Envio
                    </div>
                    <h2 className="mt-2 text-2xl font-semibold tracking-[-0.03em] text-[#34204f]">
                      Adicione uma planilha de origem
                    </h2>
                  </div>
                  <div className="rounded-full bg-[#45246b]/8 p-3 text-[#45246b]">
                    <UploadCloud className="h-5 w-5" />
                  </div>
                </div>

                <div className="space-y-4">
                  <div className="rounded-2xl border border-[#eee6dc] bg-[#fcfbf9] p-4">
                    <div className="text-sm font-medium text-[#34204f]">Formato esperado</div>
                    <p className="mt-2 text-sm leading-6 text-slate-600">
                      Qualquer arquivo `.xlsx` ou `.xls` com identificadores NCT em qualquer parte da planilha.
                    </p>
                  </div>

                  <div className="grid grid-cols-2 gap-3">
                    <div className="rounded-2xl border border-[#eee6dc] bg-white p-4">
                      <div className="text-[11px] font-semibold uppercase tracking-[0.22em] text-slate-500">Salvos</div>
                      <div className="mt-3 text-3xl font-semibold tracking-[-0.04em] text-[#34204f]">{savedStudies.length}</div>
                    </div>
                    <div className="rounded-2xl border border-[#eee6dc] bg-white p-4">
                      <div className="text-[11px] font-semibold uppercase tracking-[0.22em] text-slate-500">Ao vivo</div>
                      <div className="mt-3 text-3xl font-semibold tracking-[-0.04em] text-[#34204f]">{studies.length}</div>
                    </div>
                  </div>
                </div>

                <Button
                  size="lg"
                  onClick={() => fileInputRef.current?.click()}
                  disabled={isUploading || isFetching}
                  className="avanzare-btn mt-6 h-12 w-full justify-between rounded-2xl px-5 text-sm font-semibold text-white"
                >
                  <span className="flex items-center gap-2">
                    {isUploading ? <Loader2 className="h-4 w-4 animate-spin" /> : <UploadCloud className="h-4 w-4" />}
                    {isUploading ? 'Processando planilha...' : hasData ? 'Enviar outra planilha' : 'Enviar planilha'}
                  </span>
                  <ArrowRight className="h-4 w-4" />
                </Button>
              </div>
            </div>
          </div>
        </section>

        <section className="avanzare-panel avanzare-rise p-6 md:p-8" style={{ animationDelay: '90ms' }}>
          <form
            className="space-y-5"
            onSubmit={event => {
              event.preventDefault();
              void runSearch();
            }}
          >
            <div className="flex flex-col gap-2 md:flex-row md:items-end md:justify-between">
              <div>
                <div className="text-[11px] font-semibold uppercase tracking-[0.22em] text-[#8d5f37]">Busca ao vivo</div>
                <h2 className="mt-2 text-2xl font-semibold tracking-[-0.03em] text-[#34204f]">Encontre estudos sem reenviar a planilha</h2>
                <p className="mt-2 max-w-2xl text-sm leading-6 text-slate-600">
                  A tabela não muda. O que muda é a origem: ClinicalTrials.gov devolve o estudo e todos os contatos com e-mail ou telefone, com a sede ao lado.
                </p>
              </div>
              {searchTotal != null && (
                <div className="text-sm font-medium text-[#45246b]">{searchTotal} estudos na API</div>
              )}
            </div>
            <div className="grid gap-3 md:grid-cols-2 xl:grid-cols-4">
              <div className="space-y-2">
                <span className="text-[11px] font-semibold uppercase tracking-[0.18em] text-slate-500">Condição</span>
                <OptionPicker
                  options={CONDITIONS}
                  selected={searchForm.condition}
                  onChange={condition => setSearchForm(prev => ({ ...prev, condition }))}
                />
              </div>
              <div className="space-y-2">
                <span className="text-[11px] font-semibold uppercase tracking-[0.18em] text-slate-500">Droga</span>
                <OptionPicker
                  options={DRUGS}
                  selected={searchForm.drug}
                  allowAll
                  onChange={drug => setSearchForm(prev => ({ ...prev, drug }))}
                />
              </div>
              <div className="space-y-2">
                <span className="text-[11px] font-semibold uppercase tracking-[0.18em] text-slate-500">Países</span>
                <OptionPicker
                  options={COUNTRIES}
                  selected={searchForm.countries}
                  allowCustom={false}
                  allowAll
                  onChange={countries => setSearchForm(prev => ({ ...prev, countries }))}
                />
              </div>
              <div className="space-y-2">
                <span className="text-[11px] font-semibold uppercase tracking-[0.18em] text-slate-500">Patrocinadores</span>
                <OptionPicker
                  options={SPONSORS}
                  selected={searchForm.sponsors}
                  onChange={sponsors => setSearchForm(prev => ({ ...prev, sponsors }))}
                />
              </div>
              {SEARCH_DROPDOWNS.map(filter => (
                <label key={filter.key} className="space-y-2">
                  <span className="text-[11px] font-semibold uppercase tracking-[0.18em] text-slate-500">{filter.label}</span>
                  <Select
                    value={searchForm[filter.key]}
                    onValueChange={value => setSearchForm(prev => ({ ...prev, [filter.key]: value }))}
                  >
                    <SelectTrigger>
                      <SelectValue />
                    </SelectTrigger>
                    <SelectContent>
                      {filter.options.map(option => (
                        <SelectItem key={`${filter.key}-${option.value}`} value={option.value}>{option.label}</SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </label>
              ))}
            </div>
            <div className="flex flex-wrap items-end gap-3">
              <div className="space-y-2">
                <span className="text-[11px] font-semibold uppercase tracking-[0.18em] text-slate-500">Por vez</span>
                <Select value={String(pageSize)} onValueChange={value => setPageSize(Number(value))}>
                  <SelectTrigger className="h-11 w-[120px] rounded-2xl">
                    <SelectValue />
                  </SelectTrigger>
                  <SelectContent>
                    {PAGE_SIZES.map(size => (
                      <SelectItem key={size} value={String(size)}>{size}</SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
              <Button
                type="submit"
                disabled={isSearching || isFetching}
                className="avanzare-btn h-11 rounded-2xl px-6 text-white"
              >
                {isSearching ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <Search className="mr-2 h-4 w-4" />}
                Buscar estudos
              </Button>
            </div>
          </form>
        </section>

        {(isFetching || isRefreshing) && progress.total > 0 && (
          <Card className="overflow-hidden rounded-[26px] border border-[#eadfce] bg-[linear-gradient(135deg,rgba(69,36,107,0.05),rgba(255,255,255,0.88))] shadow-[0_16px_36px_rgba(48,26,70,0.06)]">
            <CardContent className="flex items-center gap-4 p-5">
              <Loader2 className="h-5 w-5 shrink-0 animate-spin text-[#45246b]" />
              <div className="flex-1 space-y-1">
                <div className="flex justify-between text-sm font-medium">
                  <span className="text-[#45246b]">{isRefreshing ? 'Atualizando contatos da lista' : 'Montando sua base de estudos'}: {progress.current} de {progress.total}</span>
                  <span className="text-[#45246b]/70">{Math.round((progress.current / progress.total) * 100)}%</span>
                </div>
                <Progress value={(progress.current / progress.total) * 100} className="h-2" />
              </div>
            </CardContent>
          </Card>
        )}

        {!hasData && savedStudies.length === 0 && (
          <div className="flex justify-center py-2">
            <Card className="w-full max-w-4xl overflow-hidden rounded-[30px] border border-[#e7dfd3] bg-white shadow-[0_20px_50px_rgba(48,26,70,0.06)]">
              <CardContent className="grid gap-8 p-8 md:grid-cols-[1.1fr_0.9fr] md:p-10">
                <div className="space-y-5">
                  <img
                    src="/avanzare-logo.jpg"
                    alt="Avanzare Clinical Research Solutions"
                    className="h-auto w-[190px]"
                  />
                  <div className="space-y-3">
                    <h3 className="text-3xl font-semibold tracking-[-0.04em] text-[#34204f]">
                      Comece com sua planilha.
                    </h3>
                    <p className="max-w-xl text-sm leading-7 text-slate-600">
                      Envie um arquivo Excel com identificadores NCT e revise os estudos resultantes em um espaço Avanzare mais limpo.
                    </p>
                  </div>
                  <div className="rounded-[24px] border border-[#ede1ce] bg-[#fbf6ee] p-5">
                    <div className="text-[11px] font-semibold uppercase tracking-[0.22em] text-[#8d5f37]">
                      Fluxo
                    </div>
                    <p className="mt-3 text-sm leading-7 text-slate-600">
                      Envie o arquivo, revise os dados capturados, mantenha os estudos relevantes e exporte o CSV final quando estiver pronto.
                    </p>
                  </div>
                </div>
                <div className="flex flex-col justify-center rounded-[26px] border border-[#ece3d8] bg-[#fcfbf8] p-8 text-center">
                  <div className="mx-auto mb-6 flex h-18 w-18 items-center justify-center rounded-full bg-[#45246b]/10">
                    <UploadCloud className="h-9 w-9 text-[#45246b]" />
                  </div>
                  <h3 className="text-xl font-semibold text-[#34204f]">Enviar identificadores de estudos</h3>
                  <p className="mx-auto mt-3 max-w-sm text-sm leading-6 text-slate-500">
                    Lemos a planilha inteira, detectamos os IDs NCT automaticamente e montamos um espaço limpo para revisão.
                  </p>
                  <Button size="lg" onClick={() => fileInputRef.current?.click()} disabled={isUploading} className="avanzare-btn mt-8 h-12 rounded-2xl px-8 text-white">
                    {isUploading ? <><Loader2 className="mr-2 h-4 w-4 animate-spin" /> Processando...</> : 'Selecionar arquivo Excel'}
                  </Button>
                </div>
              </CardContent>
            </Card>
          </div>
        )}

        {(hasData || savedStudies.length > 0) && (
          <section className="space-y-4">
            <div className="flex flex-col gap-3 md:flex-row md:items-end md:justify-between">
              <div>
                <p className="text-xs font-semibold uppercase tracking-[0.2em] text-[#8d5f37]">Espaço de trabalho</p>
                <h2 className="mt-2 text-2xl font-semibold tracking-[-0.03em] text-[#34204f]">Revise, selecione e exporte seus estudos</h2>
                <p className="mt-2 text-sm text-slate-500">
                  Use a aba de resultados para a saída atual da API e a aba de shortlist para a seleção da Avanzare.
                </p>
              </div>
              <div className="rounded-full border border-blue-200 bg-white px-4 py-2 font-mono text-xs font-medium text-[#0e3a6b]">
                Fonte pública • ClinicalTrials.gov • Fluxo guiado por Excel
              </div>
            </div>

            <Tabs value={activeTab} onValueChange={setActiveTab}>
              <TabsList className="h-auto rounded-[22px] border border-blue-200 bg-white p-1.5">
                <TabsTrigger value="results" className="avanzare-tab gap-2 rounded-[16px] px-5 py-3 data-[state=active]:text-white data-[state=active]:shadow-none">
                  <DatabaseZap className="h-4 w-4" />
                  Resultados ao vivo
                  {studies.length > 0 && (
                    <Badge variant="secondary" className="ml-1 bg-white/14 px-1.5 py-0 text-xs text-inherit">{studies.length}</Badge>
                  )}
                </TabsTrigger>
                <TabsTrigger value="saved" className="avanzare-tab gap-2 rounded-[16px] px-5 py-3 data-[state=active]:text-white data-[state=active]:shadow-none">
                  <List className="h-4 w-4" />
                  Lista Avanzare
                  {savedStudies.length > 0 && (
                    <Badge variant="secondary" className="ml-1 bg-white/14 px-1.5 py-0 text-xs text-inherit">{savedStudies.length}</Badge>
                  )}
                </TabsTrigger>
              </TabsList>

              <TabsContent value="results" className="mt-5">
                {studies.length > 0 ? (
                  <>
                    <div className="mb-3 flex justify-end">
                      <Button size="sm" variant="outline" onClick={() => exportCSV(studies, contactFocus)} disabled={studies.length === 0 || isFetching} className="gap-2 rounded-xl border-[#ece3d8] bg-white/90 hover:bg-white">
                        <Download className="h-4 w-4" /> Exportar CSV
                      </Button>
                    </div>
                    <StudyTable
                      studies={studies}
                      isFetching={isFetching || isSearching}
                      contactFocus={contactFocus}
                      actionSlot={(study) =>
                        savedIds.has(study.nctId) ? (
                          <Button size="sm" variant="ghost" disabled className="h-7 gap-1 px-2 text-xs text-emerald-600">
                            <BookmarkCheck className="h-3.5 w-3.5" /> Salvo
                          </Button>
                        ) : (
                          <Button
                            size="sm"
                            variant="outline"
                            className="h-7 gap-1 rounded-xl border-[#ece3d8] bg-white/90 px-2 text-xs"
                            onClick={() => saveStudy(study)}
                          >
                            <Bookmark className="h-3.5 w-3.5" /> Salvar
                          </Button>
                        )
                      }
                    />
                    {nextPageToken && (
                      <div className="mt-3 flex justify-center">
                        <Button variant="outline" onClick={() => void loadMore()} disabled={isSearching} className="rounded-xl border-[#ece3d8] bg-white">
                          {isSearching && <Loader2 className="mr-2 h-4 w-4 animate-spin" />}
                          Carregar mais{searchTotal != null ? ` (${studies.length} de ${searchTotal})` : ''}
                        </Button>
                      </div>
                    )}
                  </>
                ) : (
                  !isFetching && (
                    <Card className="mt-4 rounded-[24px] border border-dashed border-[#ddcfbd] bg-white shadow-none">
                      <CardContent className="flex flex-col items-center justify-center py-16 text-slate-500">
                        <UploadCloud className="mb-3 h-8 w-8 text-[#b48654]" />
                        Busque ao vivo ou envie uma planilha para preencher a tabela.
                      </CardContent>
                    </Card>
                  )
                )}
              </TabsContent>

              <TabsContent value="saved" className="mt-5">
                <SavedList
                  saved={savedStudies}
                  onRemove={removeStudy}
                  onExport={() => exportCSV(savedStudies)}
                  onUpdateOrigin={updateSavedOrigin}
                  onRefresh={() => void refreshSaved()}
                  isRefreshing={isRefreshing}
                />
              </TabsContent>
            </Tabs>
          </section>
        )}
      </div>
    </div>
  );
}
