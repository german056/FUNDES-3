import { useState, useEffect } from 'react';
import { Course } from './types';
import Dashboard from './components/Dashboard';
import CourseTable from './components/CourseTable';
import CourseForm from './components/CourseForm';
import StatsReports from './components/StatsReports';
import AuthModal from './components/AuthModal';
import { 
  BookOpen, Layout, PieChart, FileText, RefreshCw, 
  HelpCircle, AlertCircle, CheckCircle2, LogIn, LogOut, Lock,
  Globe, Database, Server, Cpu, FileJson, Check, Copy, ArrowRight, ShieldCheck, Download
} from 'lucide-react';
import * as xlsx from 'xlsx';

// Cursos iniciales de respaldo para cuando no existe conexión al servidor ni datos previos
const BACKUP_COURSES: Course[] = [
  {
    id: '1',
    name: 'Especialización en Big Data y Analítica',
    url: 'https://virtual.fundes.edu.co/course/view.php?id=101',
    progress: 100,
    status: 'green',
    category: 'Especializaciones',
    responsible: 'Dr. Fernando Gómez',
    notes: 'Curso completamente ajustado y operacional',
    imageUrl: 'https://images.unsplash.com/photo-1551288049-bebda4e38f71?auto=format&fit=crop&w=600&q=80'
  },
  {
    id: '2',
    name: 'Introducción al Desarrollo Frontend con React',
    url: 'https://virtual.fundes.edu.co/course/view.php?id=102',
    progress: 75,
    status: 'yellow',
    category: 'Ingeniería de Software',
    responsible: 'MSc. Elena Rostova',
    notes: 'Faltan subir algunos recursos del módulo 4, pero es funcional',
    imageUrl: 'https://images.unsplash.com/photo-1633356122544-f134324a6cee?auto=format&fit=crop&w=600&q=80'
  },
  {
    id: '3',
    name: 'Gestión de Proyectos E-Learning con Moodle',
    url: 'https://virtual.fundes.edu.co/course/view.php?id=103',
    progress: 30,
    status: 'red',
    category: 'Educación Virtual',
    responsible: 'Ing. Carlos Mendoza',
    notes: 'Enlaces rotos en exámenes finales y recursos de bienvenida inactivos',
    imageUrl: 'https://images.unsplash.com/photo-1516321318423-f06f85e504b3?auto=format&fit=crop&w=600&q=80'
  },
  {
    id: '4',
    name: 'Especialidad en Gerencia Social',
    url: 'https://virtual.fundes.edu.co/course/view.php?id=104',
    progress: 95,
    status: 'yellow',
    category: 'Administración',
    responsible: 'Dra. Patricia Ortiz',
    notes: 'Ajustes menores de diseño en el banner principal',
    imageUrl: 'https://images.unsplash.com/photo-1522071820081-009f0129c71c?auto=format&fit=crop&w=600&q=80'
  },
  {
    id: '5',
    name: 'Diseño Curricular para Entornos Virtuales',
    url: 'https://virtual.fundes.edu.co/course/view.php?id=105',
    progress: 100,
    status: 'green',
    category: 'Educación Virtual',
    responsible: 'Dra. Laura Marín',
    notes: 'Aprobado al 100% por coordinación académica',
    imageUrl: 'https://images.unsplash.com/photo-1515187029135-18ee286d815b?auto=format&fit=crop&w=600&q=80'
  }
];

// Parser de Google Sheets compatible 100% con ejecución en navegadores (para Hostinger)
const parseSpreadsheetClientSide = async (): Promise<Course[]> => {
  const url = "https://docs.google.com/spreadsheets/d/1L5jBs1ChPJ8NGbTE8OC1AbvKFZ_qMdYH/export?format=xlsx";
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`Error de red al descargar el archivo: ${response.statusText}`);
  }
  const buffer = await response.arrayBuffer();
  const workbook = xlsx.read(new Uint8Array(buffer), { type: 'array' });
  
  const coursesLoaded: Course[] = [];
  let idCounter = Date.now();

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const rawRows = xlsx.utils.sheet_to_json<any[]>(sheet, { header: 1 });
    if (rawRows.length === 0) continue;

    for (let i = 0; i < rawRows.length; i++) {
      const row = rawRows[i];
      if (!row || row.length < 2) continue;

      const rowStr = row.join(' ').toLowerCase();
      if (rowStr.includes('estandar') && rowStr.includes('no funciona') && rowStr.includes('funcional')) {
        continue;
      }

      const urlIndex = row.findIndex((cell: any) => typeof cell === 'string' && (cell.includes('http://') || cell.includes('https://')));
      
      if (urlIndex !== -1) {
        const courseUrl = row[urlIndex];
        let courseName = '';
        let courseCategory = sheetName;
        let courseResponsible = '';
        let courseStatus: 'green' | 'yellow' | 'red' = 'green';
        let courseProgress = 100;
        let courseNotes = '';

        for (let j = 0; j < row.length; j++) {
          const cellVal = String(row[j] || '').toUpperCase().trim();
          if (cellVal === 'ESTANDAR' || cellVal === 'FUNCIONAL COMPLETO' || cellVal === 'ESTÁNDAR') {
            courseStatus = 'green';
            courseProgress = 100;
          } else if (cellVal === 'FUNCIONAL' || cellVal === 'FUNCIONAL CON MEJORAS' || cellVal === 'CON MEJORAS' || cellVal === 'MEJORAR') {
            courseStatus = 'yellow';
            courseProgress = 75;
          } else if (cellVal === 'NO FUNCIONA' || cellVal === 'CRÍTICO' || cellVal === 'CRITICO' || cellVal === 'MALO' || cellVal === 'FALLA') {
            courseStatus = 'red';
            courseProgress = 20;
          }
        }

        const textCellIndices = row
          .map((cell: any, idx: number) => ({val: cell, idx}))
          .filter((item: any) => typeof item.val === 'string' && item.val.trim().length > 3 && item.idx !== urlIndex);

        if (textCellIndices.length > 0) {
          courseName = textCellIndices[0].val;
          if (textCellIndices.length > 1) {
            const secondVal = textCellIndices[1].val;
            if (secondVal.includes(',') || secondVal.split(' ').length >= 2) {
              courseResponsible = secondVal;
            } else {
              courseCategory = secondVal;
            }
          }
          if (textCellIndices.length > 2) {
            const thirdVal = textCellIndices[2].val;
            if (!courseResponsible) {
              courseResponsible = thirdVal;
            } else {
              courseNotes = thirdVal;
            }
          }
        }

        if (!courseName) {
          courseName = `Curso Virtual #${idCounter}`;
        }

        const textToAnalyze = `${courseName} ${courseCategory}`.toLowerCase();
        let parsedImageUrl = 'https://images.unsplash.com/photo-1516321318423-f06f85e504b3?auto=format&fit=crop&w=600&q=80';
        if (textToAnalyze.includes('data') || textToAnalyze.includes('analítica') || textToAnalyze.includes('big data') || textToAnalyze.includes('software') || textToAnalyze.includes('react') || textToAnalyze.includes('frontend')) {
          parsedImageUrl = 'https://images.unsplash.com/photo-1551288049-bebda4e38f71?auto=format&fit=crop&w=600&q=80';
        } else if (textToAnalyze.includes('gerencia') || textToAnalyze.includes('proyecto') || textToAnalyze.includes('administración')) {
          parsedImageUrl = 'https://images.unsplash.com/photo-1522071820081-009f0129c71c?auto=format&fit=crop&w=600&q=80';
        } else if (textToAnalyze.includes('pedagogía') || textToAnalyze.includes('diseño curricular') || textToAnalyze.includes('docencia') || textToAnalyze.includes('virtual')) {
          parsedImageUrl = 'https://images.unsplash.com/photo-1515187029135-18ee286d815b?auto=format&fit=crop&w=600&q=80';
        }

        coursesLoaded.push({
          id: String(idCounter++),
          name: courseName.trim(),
          url: courseUrl,
          progress: courseProgress,
          status: courseStatus,
          category: courseCategory.trim(),
          responsible: courseResponsible.trim() || 'Coordinación Virtual',
          notes: courseNotes.trim() || 'Sincronizado desde plantilla',
          imageUrl: parsedImageUrl
        });
      }
    }
  }

  if (coursesLoaded.length > 0) {
    return coursesLoaded;
  } else {
    // Intentar análisis alternativo si no se encontraron celdas de tipo URL explícitas
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const jsonData = xlsx.utils.sheet_to_json<any[]>(worksheet, { header: 1 });
    const fallbackList = parseSpreadsheetGenericFallbackClientSide(jsonData);
    if (fallbackList) return fallbackList;
    throw new Error("No se encontraron registros de cursos en la plantilla.");
  }
};

const parseSpreadsheetGenericFallbackClientSide = (grid: any[][]): Course[] | null => {
  if (!grid || grid.length < 2) return null;
  const list: Course[] = [];
  let id = Date.now();
  
  for (let r = 0; r < grid.length; r++) {
    const row = grid[r];
    if (!row || row.length < 3) continue;
    
    const nameCandidate = row.find(c => typeof c === 'string' && c.length > 6 && !c.includes('http') && !c.includes('@'));
    const urlCandidate = row.find(c => typeof c === 'string' && c.startsWith('http'));
    
    if (nameCandidate && urlCandidate) {
      let category = 'Virtualización';
      const categoryCandidate = row.find(c => typeof c === 'string' && c.length > 3 && c !== nameCandidate && c !== urlCandidate && !c.includes('@'));
      if (categoryCandidate) category = categoryCandidate;
      
      const responsibleCandidate = row.find(c => typeof c === 'string' && c.includes('@')) || 'Docente Asignado';
      
      let status: 'green' | 'yellow' | 'red' = 'green';
      let progress = 100;
      const progressCandidate = row.find(c => typeof c === 'number');
      if (progressCandidate) {
        progress = Math.min(100, Math.max(0, progressCandidate));
        status = progress === 100 ? 'green' : progress >= 60 ? 'yellow' : 'red';
      }

      const fallbackAnalyze = `${nameCandidate} ${category}`.toLowerCase();
      let fallbackImageUrl = 'https://images.unsplash.com/photo-1516321318423-f06f85e504b3?auto=format&fit=crop&w=600&q=80';
      if (fallbackAnalyze.includes('data') || fallbackAnalyze.includes('analítica') || fallbackAnalyze.includes('software') || fallbackAnalyze.includes('react')) {
        fallbackImageUrl = 'https://images.unsplash.com/photo-1551288049-bebda4e38f71?auto=format&fit=crop&w=600&q=80';
      } else if (fallbackAnalyze.includes('gerencia') || fallbackAnalyze.includes('proyecto') || fallbackAnalyze.includes('administración')) {
        fallbackImageUrl = 'https://images.unsplash.com/photo-1522071820081-009f0129c71c?auto=format&fit=crop&w=600&q=80';
      }

      list.push({
        id: String(id++),
        name: nameCandidate,
        url: urlCandidate,
        progress,
        status,
        category,
        responsible: responsibleCandidate,
        notes: 'Importado de plantilla localmente',
        imageUrl: fallbackImageUrl
      });
    }
  }
  return list.length > 0 ? list : null;
};

export default function App() {
  const [courses, setCourses] = useState<Course[]>([]);
  const [activeTab, setActiveTab] = useState<'dashboard' | 'table' | 'reports'>('dashboard');
  const [isFormOpen, setIsFormOpen] = useState(false);
  const [selectedCourse, setSelectedCourse] = useState<Course | null>(null);
  const [categories, setCategories] = useState<string[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [isSyncing, setIsSyncing] = useState(false);
  const [notifications, setNotifications] = useState<{ type: 'success' | 'error'; msg: string } | null>(null);

  // Estados de control para Despliegue de Hostinger y Almacenamiento Autónomo
  const [isLocalOnly, setIsLocalOnly] = useState<boolean>(() => {
    return localStorage.getItem('fundes_hostinger_local_mode') === 'true';
  });
  const [isHostingerModalOpen, setIsHostingerModalOpen] = useState<boolean>(false);

  // Authentication states with Local Storage persistence
  const [isAuthenticated, setIsAuthenticated] = useState<boolean>(() => {
    return localStorage.getItem('fundes_authenticated') === 'true';
  });
  const [isAuthOpen, setIsAuthOpen] = useState<boolean>(false);
  const [pendingAction, setPendingAction] = useState<(() => void) | null>(null);

  // Interceptor helper to run actions only if authenticated
  const executeWithAuth = (action: () => void) => {
    if (isAuthenticated) {
      action();
    } else {
      setPendingAction(() => action);
      setIsAuthOpen(true);
      triggerNotification('error', 'Identificación requerida para realizar modificaciones (Usuario/Clave: 1234).');
    }
  };

  const handleAuthSuccess = () => {
    setIsAuthenticated(true);
    localStorage.setItem('fundes_authenticated', 'true');
    setIsAuthOpen(false);
    triggerNotification('success', '¡Autenticación de Coordinador exitosa!');
    
    // Execute the action that triggered the login modal, if any
    if (pendingAction) {
      pendingAction();
      setPendingAction(null);
    }
  };

  // Load courses on startup
  useEffect(() => {
    fetchCourses();
  }, []);

  const triggerNotification = (type: 'success' | 'error', msg: string) => {
    setNotifications({ type, msg });
    setTimeout(() => {
      setNotifications(null);
    }, 5000);
  };

  const fetchCourses = async () => {
    setIsLoading(true);
    
    // Si el usuario forzó manualmente el modo local de Hostinger
    if (localStorage.getItem('fundes_hostinger_local_mode') === 'true') {
      loadFromBrowserLocalStorage();
      return;
    }

    try {
      // Intentar conectar con el back-end de Node.js en desarrollo (timeout corto de 2.5s)
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 2500);
      
      const response = await fetch('/api/courses', { signal: controller.signal });
      clearTimeout(timeoutId);
      
      if (!response.ok) throw new Error('Servidor retornó código erróneo');
      const data = await response.json();
      setCourses(data.courses || []);
      setIsLocalOnly(false);
      
      const cats = Array.from(new Set((data.courses || []).map((c: Course) => c.category))) as string[];
      setCategories(cats.length ? cats : ['Especializaciones', 'Educación Virtual', 'Ingeniería']);
    } catch (error: any) {
      console.warn("Fallo al conectar con el servidor Express v1. Activando automáticamente el Modo Autónomo de Hostinger (LocalStorage):", error);
      setIsLocalOnly(true);
      loadFromBrowserLocalStorage();
    } finally {
      setIsLoading(false);
    }
  };

  const loadFromBrowserLocalStorage = () => {
    const localDataRaw = localStorage.getItem('fundes_courses');
    if (localDataRaw) {
      try {
        const parsed = JSON.parse(localDataRaw) as Course[];
        setCourses(parsed);
        const cats = Array.from(new Set(parsed.map((c: Course) => c.category))) as string[];
        setCategories(cats.length ? cats : ['Especializaciones', 'Educación Virtual', 'Ingeniería']);
      } catch (e) {
        console.error("Fallo al leer datos del navegador, reiniciando:", e);
        setCourses(BACKUP_COURSES);
        localStorage.setItem('fundes_courses', JSON.stringify(BACKUP_COURSES));
        setCategories(['Especializaciones', 'Educación Virtual', 'Ingeniería']);
      }
    } else {
      setCourses(BACKUP_COURSES);
      localStorage.setItem('fundes_courses', JSON.stringify(BACKUP_COURSES));
      setCategories(['Especializaciones', 'Educación Virtual', 'Ingeniería']);
    }
  };

  // Synchronize/pull courses directly from google sheets template URL (indispensable)
  const handleSyncSheets = async () => {
    setIsSyncing(true);
    triggerNotification('success', 'Sincronizando plantilla especializaciones virtuales de Google Sheets...');
    
    // Si estamos en modo de Hostinger Autónomo (LocalStorage), sincronizar 100% en cliente
    if (isLocalOnly) {
      try {
        const sheetData = await parseSpreadsheetClientSide();
        setCourses(sheetData);
        localStorage.setItem('fundes_courses', JSON.stringify(sheetData));
        
        const cats = Array.from(new Set(sheetData.map((c: Course) => c.category))) as string[];
        setCategories(cats.length ? cats : ['Especializaciones', 'Educación Virtual', 'Ingeniería']);
        
        triggerNotification('success', `¡Sincronización Hostinger Exitosa! Se cargaron ${sheetData.length} cursos virtuales directamente desde tu navegador.`);
      } catch (error: any) {
        console.error("Fallo de sincronización directa client-side:", error);
        triggerNotification('error', `Error de CORS o Red al sincronizar: ${error.message || 'La plantilla está sin acceso público '}`);
      } finally {
        setIsSyncing(false);
      }
      return;
    }

    // Modo estándar si está encendido el servidor Node.js
    try {
      const response = await fetch('/api/courses/sync');
      if (!response.ok) throw new Error('Error al sincronizar con Google Sheets');
      const data = await response.json();
      setCourses(data.courses || []);
      
      const cats = Array.from(new Set((data.courses || []).map((c: Course) => c.category))) as string[];
      setCategories(cats.length ? cats : ['Especializaciones', 'Educación Virtual']);
      
      triggerNotification('success', `¡Sincronización completa! Se cargaron ${data.count} cursos virtuales.`);
    } catch (error: any) {
      console.warn("Fallo al llamar al API de sincronización, intentando modo cliente directo de respaldo:", error);
      try {
        const sheetData = await parseSpreadsheetClientSide();
        setCourses(sheetData);
        localStorage.setItem('fundes_courses', JSON.stringify(sheetData));
        setIsLocalOnly(true);
        localStorage.setItem('fundes_hostinger_local_mode', 'true');
        
        const cats = Array.from(new Set(sheetData.map((c: Course) => c.category))) as string[];
        setCategories(cats.length ? cats : ['Especializaciones', 'Educación Virtual', 'Ingeniería']);
        
        triggerNotification('success', 'Sincronización de respaldo directa completada correctamente.');
      } catch (err2) {
        triggerNotification('error', 'Error de red o CORS al intentar conectar con plantillas de Google.');
      }
    } finally {
      setIsSyncing(false);
    }
  };

  // Create or Update course virtual
  const handleSaveCourse = async (courseData: Omit<Course, 'id'> & { id?: string }) => {
    if (isLocalOnly) {
      try {
        let updatedCourses: Course[];
        if (courseData.id) {
          // Editar curso existente
          updatedCourses = courses.map(c => c.id === courseData.id ? { ...c, ...courseData } as Course : c);
          triggerNotification('success', 'Curso actualizado exitosamente en el navegador.');
        } else {
          // Crear un nuevo curso
          const newCourse: Course = {
            ...courseData,
            id: String(Date.now()),
            imageUrl: courseData.imageUrl || 'https://images.unsplash.com/photo-1516321318423-f06f85e504b3?auto=format&fit=crop&w=600&q=80'
          };
          updatedCourses = [...courses, newCourse];
          triggerNotification('success', 'Curso registrado de manera exitosa en el navegador.');
        }
        setCourses(updatedCourses);
        localStorage.setItem('fundes_courses', JSON.stringify(updatedCourses));
        
        const cats = Array.from(new Set(updatedCourses.map((c: Course) => c.category))) as string[];
        setCategories(cats.length ? cats : ['Especializaciones', 'Educación Virtual']);
        
        setIsFormOpen(false);
        setSelectedCourse(null);
      } catch (err) {
        triggerNotification('error', 'No se pudo guardar la información localmente.');
      }
      return;
    }

    try {
      let response;
      if (courseData.id) {
        // Edit existing course
        response = await fetch(`/api/courses/${courseData.id}`, {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(courseData)
        });
      } else {
        // Create new course
        response = await fetch('/api/courses', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(courseData)
        });
      }

      if (!response.ok) throw new Error('Error al guardar el curso académico');
      await fetchCourses();
      setIsFormOpen(false);
      setSelectedCourse(null);
      triggerNotification('success', courseData.id ? 'Curso actualizado exitosamente.' : 'Curso registrado de manera exitosa.');
    } catch (error: any) {
      console.error(error);
      triggerNotification('error', 'No se pudo guardar la información en la base de datos de red.');
    }
  };

  // Delete course
  const handleDeleteCourse = async (id: string) => {
    if (isLocalOnly) {
      try {
        const updatedCourses = courses.filter(c => c.id !== id);
        setCourses(updatedCourses);
        localStorage.setItem('fundes_courses', JSON.stringify(updatedCourses));
        
        const cats = Array.from(new Set(updatedCourses.map((c: Course) => c.category))) as string[];
        setCategories(cats.length ? cats : ['Especializaciones', 'Educación Virtual']);
        
        triggerNotification('success', 'Curso académico eliminado localmente del navegador.');
      } catch (err) {
        triggerNotification('error', 'Fallo al procesar la eliminación local.');
      }
      return;
    }

    try {
      const response = await fetch(`/api/courses/${id}`, {
        method: 'DELETE'
      });
      if (!response.ok) throw new Error('Error al eliminar curso');
      await fetchCourses();
      triggerNotification('success', 'Curso académico eliminado correctamente.');
    } catch (error: any) {
      console.error(error);
      triggerNotification('error', 'Fallo al procesar la eliminación en el servidor.');
    }
  };

  const handleEditClick = (course: Course) => {
    setSelectedCourse(course);
    setIsFormOpen(true);
  };

  const handleCreateClick = () => {
    setSelectedCourse(null);
    setIsFormOpen(true);
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-800 flex flex-col font-sans">
      
      {/* Top Header */}
      <header className="bg-white border-b border-slate-100 sticky top-0 z-40 print:hidden shrink-0">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-4 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
          
          {/* Logo & Platform Name */}
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 bg-indigo-600 rounded-2xl flex items-center justify-center text-white font-extrabold text-lg shadow-md shadow-indigo-600/20">
              FV
            </div>
            <div>
              <h1 className="text-lg font-extrabold text-slate-800 tracking-tight flex items-center gap-2">
                Gestión de Cursos Virtuales
                <span className="text-[10px] font-bold tracking-widest uppercase py-0.5 px-2 bg-indigo-50 text-indigo-700 rounded-md border border-indigo-100/30">
                  Fundes
                </span>
              </h1>
              <p className="text-xs text-slate-400 font-medium">FUNDES, FUNDACIÓN DE ESTUDIOS SUPERIORES MONSEÑOR ABRAHAM ESCUDERO MONTOYA</p>
            </div>
          </div>

          {/* Sync Sheets Controls & Authentication Status */}
          <div className="flex flex-wrap items-center gap-2.5">
            {/* Indicador de Estado de Datos Híbrido */}
            {isLocalOnly ? (
              <div 
                onClick={() => setIsHostingerModalOpen(true)}
                className="flex items-center gap-1.5 bg-amber-50 text-amber-800 border border-amber-200/60 px-3 py-1.5 rounded-xl text-[10px] font-bold cursor-pointer hover:bg-amber-100/80 transition-all select-none"
                title="Haga clic para ver detalles de Hostinger"
              >
                <div className="w-1.5 h-1.5 rounded-full bg-amber-500 animate-pulse shrink-0" />
                <Database size={11} className="text-amber-600" />
                <span>Autónomo (Hostinger Local)</span>
              </div>
            ) : (
              <div 
                onClick={() => setIsHostingerModalOpen(true)}
                className="flex items-center gap-1.5 bg-indigo-50 text-indigo-800 border border-indigo-100 px-3 py-1.5 rounded-xl text-[10px] font-bold cursor-pointer hover:bg-indigo-100/60 transition-all select-none"
                title="Haga clic para configuraciones de despliegue"
              >
                <div className="w-1.5 h-1.5 rounded-full bg-indigo-500 animate-pulse shrink-0" />
                <Server size={11} className="text-indigo-600" />
                <span>Servidor Express (Dev)</span>
              </div>
            )}

            {isAuthenticated ? (
              <div className="flex items-center gap-2 bg-slate-100/70 border border-slate-200/80 px-3.5 py-2 rounded-xl">
                <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse shrink-0" />
                <span className="text-[11px] font-bold text-slate-700">Coordinador Principal</span>
                <button
                  type="button"
                  onClick={() => {
                    localStorage.removeItem('fundes_authenticated');
                    setIsAuthenticated(false);
                    triggerNotification('success', 'Sesión de Coordinador cerrada exitosamente.');
                  }}
                  className="ml-1 text-slate-400 hover:text-rose-600 hover:bg-white p-1 rounded-lg transition-all border border-transparent hover:border-slate-150 cursor-pointer"
                  title="Cerrar sesión"
                >
                  <LogOut size={12} />
                </button>
              </div>
            ) : (
              <button
                type="button"
                onClick={() => {
                  setPendingAction(null);
                  setIsAuthOpen(true);
                }}
                className="px-3.5 py-2 rounded-xl text-slate-700 bg-slate-100 hover:bg-indigo-50 hover:text-indigo-700/90 border border-slate-200 hover:border-indigo-100/50 font-bold text-[11px] transition-all flex items-center gap-1.5 cursor-pointer"
              >
                <LogIn size={12} />
                <span>Iniciar Sesión</span>
              </button>
            )}

            <button
              onClick={() => executeWithAuth(handleSyncSheets)}
              disabled={isSyncing}
              className={`px-3.5 py-2 rounded-xl text-slate-700 bg-white border border-slate-200 hover:bg-slate-50 font-semibold text-xs transition-all flex items-center gap-2 active:scale-95 disabled:opacity-60 disabled:pointer-events-none cursor-pointer`}
            >
              <RefreshCw size={13} className={isSyncing ? 'animate-spin text-indigo-500' : 'text-slate-400'} />
              <span>Sincronizar Excel</span>
            </button>

            <button
              onClick={() => setIsHostingerModalOpen(true)}
              className="px-3.5 py-2 rounded-xl text-white bg-indigo-600 hover:bg-indigo-700 font-bold text-[11px] transition-all flex items-center gap-1.5 shadow-sm shadow-indigo-600/20 hover:shadow-indigo-600/30 active:scale-95 cursor-pointer"
            >
              <Globe size={12} className="animate-pulse text-indigo-100" />
              <span>Desplegar en Hostinger</span>
            </button>
          </div>

        </div>
      </header>

      {/* Persistent Notification system */}
      {notifications && (
        <div className="fixed top-18 right-6 z-50 print:hidden transition-all duration-300">
          <div className={`p-4 rounded-xl border shadow-lg flex items-start gap-3 max-w-sm ${notifications.type === 'success' ? 'bg-emerald-50 border-emerald-100 text-emerald-800' : 'bg-rose-50 border-rose-100 text-rose-800'}`}>
            <div className="shrink-0 mt-0.5">
              {notifications.type === 'success' ? <CheckCircle2 size={16} /> : <AlertCircle size={16} />}
            </div>
            <div>
              <p className="text-xs font-bold uppercase tracking-wider">{notifications.type === 'success' ? 'Éxito' : 'Mensaje o Alerta'}</p>
              <p className="text-xs mt-0.5 font-medium leading-relaxed">{notifications.msg}</p>
            </div>
          </div>
        </div>
      )}

      {/* Navigation Sub-Tabs bar */}
      <nav className="bg-white border-b border-slate-100 py-3 px-4 sm:px-6 lg:px-8 print:hidden shrink-0">
        <div className="max-w-7xl mx-auto flex items-center gap-2">
          
          {/* Tab 1: Dashboard Control Panel */}
          <button
            onClick={() => setActiveTab('dashboard')}
            className={`px-4 py-2 rounded-xl text-xs font-bold tracking-wide uppercase transition-all flex items-center gap-1.5 ${activeTab === 'dashboard' ? 'bg-indigo-600 text-white shadow-sm shadow-indigo-600/15' : 'text-slate-500 hover:text-slate-800 hover:bg-slate-50'}`}
          >
            <PieChart size={13} /> Panel de Control
          </button>

          {/* Tab 2: Admin Table list */}
          <button
            onClick={() => setActiveTab('table')}
            className={`px-4 py-2 rounded-xl text-xs font-bold tracking-wide uppercase transition-all flex items-center gap-1.5 ${activeTab === 'table' ? 'bg-indigo-600 text-white shadow-sm shadow-indigo-600/15' : 'text-slate-500 hover:text-slate-800 hover:bg-slate-50'}`}
          >
            <Layout size={13} /> Ventana Principal (Administración)
          </button>

          {/* Tab 3: Detailed Report Audit printable sheets */}
          <button
            onClick={() => setActiveTab('reports')}
            className={`px-4 py-2 rounded-xl text-xs font-bold tracking-wide uppercase transition-all flex items-center gap-1.5 ${activeTab === 'reports' ? 'bg-indigo-600 text-white shadow-sm shadow-indigo-600/15' : 'text-slate-500 hover:text-slate-800 hover:bg-slate-50'}`}
          >
            <FileText size={13} /> Reportes Académicos
          </button>

        </div>
      </nav>

      {/* Main Content Area */}
      <main className="flex-1 max-w-7xl w-full mx-auto p-4 sm:p-6 lg:p-8 overflow-y-auto">
        
        {isLoading ? (
          <div className="h-96 flex flex-col items-center justify-center gap-3">
            <div className="w-10 h-10 border-3 border-indigo-600/20 border-t-indigo-600 rounded-full animate-spin" />
            <span className="text-xs text-slate-500 font-medium">Cargando base de datos de cursos...</span>
          </div>
        ) : (
          <div className="transition-all duration-300">
            {activeTab === 'dashboard' && <Dashboard courses={courses} />}
            {activeTab === 'table' && (
              <CourseTable 
                courses={courses} 
                categories={categories}
                onAddCourse={() => executeWithAuth(handleCreateClick)}
                onEditCourse={(course) => executeWithAuth(() => handleEditClick(course))}
                onDeleteCourse={(id) => executeWithAuth(() => handleDeleteCourse(id))}
              />
            )}
            {activeTab === 'reports' && <StatsReports courses={courses} />}
          </div>
        )}

      </main>

      {/* Footer */}
      <footer className="py-4 border-t border-slate-100 bg-white text-center text-xs text-slate-400 font-medium print:hidden shrink-0 mt-auto">
        <p>© 2026 Plataforma Especializaciones Virtuales Fundes. Todos los derechos reservados.</p>
      </footer>

      {/* Modal - Course Creator/Editor */}
      {isFormOpen && (
        <CourseForm
          course={selectedCourse}
          categories={categories}
          onSave={(courseData) => executeWithAuth(() => handleSaveCourse(courseData))}
          onCancel={() => {
            setIsFormOpen(false);
            setSelectedCourse(null);
          }}
        />
      )}

      {/* Modal - Auth Verification */}
      {isAuthOpen && (
        <AuthModal
          onClose={() => {
            setIsAuthOpen(false);
            setPendingAction(null);
          }}
          onSuccess={handleAuthSuccess}
        />
      )}

      {/* Modal - Asistente de Despliegue en Hostinger */}
      {isHostingerModalOpen && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4 overflow-y-auto">
          <div className="bg-white rounded-3xl border border-slate-100 shadow-2xl max-w-2xl w-full max-h-[92vh] flex flex-col overflow-hidden animate-in fade-in zoom-in-95 duration-200">
            
            {/* Header del Modal */}
            <div className="bg-gradient-to-r from-slate-900 to-indigo-950 p-6 text-white shrink-0">
              <div className="flex items-center justify-between">
                <div className="flex items-center gap-3">
                  <div className="w-10 h-10 bg-indigo-600 rounded-xl flex items-center justify-center text-white shadow-md">
                    <Globe size={18} className="animate-spin text-white" />
                  </div>
                  <div>
                    <h2 className="text-base font-extrabold tracking-tight">Centro de Despliegue en Hostinger</h2>
                    <p className="text-[11px] text-slate-300 font-medium">Arquitectura optimizada para servirse como SPA autónoma libre de servidores Node.js externos.</p>
                  </div>
                </div>
                <button 
                  onClick={() => setIsHostingerModalOpen(false)}
                  className="w-8 h-8 rounded-full bg-white/10 flex items-center justify-center text-white/80 hover:bg-white/20 transition-all cursor-pointer font-bold text-sm"
                >
                  ✕
                </button>
              </div>
            </div>

            {/* Contenido del Modal */}
            <div className="p-6 overflow-y-auto flex-1 space-y-5 text-slate-700">
              
              {/* Bloque Informativo del Estado Híbrido */}
              <div className="bg-slate-50 border border-slate-200/60 rounded-2xl p-4">
                <h3 className="text-xs font-bold uppercase tracking-wider text-slate-800 flex items-center gap-1.5">
                  <ShieldCheck size={14} className="text-emerald-500" />
                  Estado del Motor de Datos Híbrido
                </h3>
                <p className="text-xs text-slate-500 mt-1 leading-relaxed">
                  Para máxima polivalencia y ahorro de costes, hemos programado un **Motor Dúo**. Si la app detecta que no hay un servidor Express corriendo (p.ej. al subirla a Hostinger compartido), se conmuta automáticamente a usar **almacenamiento en el navegador** (LocalStorage) y sincroniza las planillas de Google directamente del lado del navegador.
                </p>

                {/* Switch de Simulación */}
                <div className="mt-3.5 flex flex-wrap items-center justify-between gap-4 py-2 border-t border-slate-150/60">
                  <div className="text-xs">
                    <span className="font-bold block">Modo forzado para depuración local:</span>
                    <span className="text-slate-400">Permite simular el comportamiento de Hostinger ahora mismo.</span>
                  </div>
                  <button
                    onClick={() => {
                      const previouslyLocal = localStorage.getItem('fundes_hostinger_local_mode') === 'true';
                      if (previouslyLocal) {
                        localStorage.removeItem('fundes_hostinger_local_mode');
                      } else {
                        localStorage.setItem('fundes_hostinger_local_mode', 'true');
                      }
                      window.location.reload();
                    }}
                    className={`px-3.5 py-1.5 rounded-xl font-bold text-xs transition-all ${isLocalOnly ? 'bg-amber-100 text-amber-800 hover:bg-amber-200' : 'bg-slate-200 text-slate-700 hover:bg-slate-300'} cursor-pointer`}
                  >
                    {isLocalOnly ? 'Desactivar Modo Hostinger' : 'Simular Modo Hostinger'}
                  </button>
                </div>
              </div>

              {/* Pestañas / Guía Paso a Paso */}
              <div>
                <h3 className="text-xs font-bold uppercase tracking-wider text-slate-800 flex items-center gap-1.5 mb-2.5">
                  <Cpu size={14} className="text-indigo-600" />
                  Instrucciones de Carga en Hostinger (Fácil y Gratis)
                </h3>
                
                <div className="space-y-3">
                  <div className="flex gap-3 text-xs leading-relaxed">
                    <div className="w-5 h-5 rounded-full bg-indigo-50 border border-indigo-200 text-indigo-700 font-bold flex items-center justify-center shrink-0">1</div>
                    <div>
                      <p className="font-bold text-slate-800">Compilar el proyecto de React</p>
                      <p className="text-slate-500">Ejecuta el script <code className="bg-slate-100 px-1 py-0.5 rounded text-indigo-600 font-mono">npm run build</code> en tu terminal. Esto empaquetará todas las vistas y el lector Excel <code className="bg-slate-100 px-1 py-0.5 rounded text-slate-600 font-mono">xlsx</code> en un directorio estático ligero optimizado llamado <code className="bg-slate-100 px-1 py-0.5 rounded font-bold">dist/</code>.</p>
                    </div>
                  </div>

                  <div className="flex gap-3 text-xs leading-relaxed">
                    <div className="w-5 h-5 rounded-full bg-indigo-50 border border-indigo-200 text-indigo-700 font-bold flex items-center justify-center shrink-0">2</div>
                    <div>
                      <p className="font-bold text-slate-800">Copia de Seguridad .htaccess Incluida</p>
                      <p className="text-slate-500">Hemos pre-configurado un archivo especial <code className="font-mono">.htaccess</code> en la carpeta de distribución <code className="font-mono">public/</code>. Cuando compiles el software, este configurador se agregará automáticamente al compilado para evitar caídas de página 404 al recargar el navegador en Hostinger.</p>
                    </div>
                  </div>

                  <div className="flex gap-3 text-xs leading-relaxed">
                    <div className="w-5 h-5 rounded-full bg-indigo-50 border border-indigo-200 text-indigo-700 font-bold flex items-center justify-center shrink-0">3</div>
                    <div>
                      <p className="font-bold text-slate-800">Sube los archivos compilados a "public_html"</p>
                      <p className="text-slate-500">Entra al panel de control hPanel de Hostinger, despliega el <strong>Administrador de Archivos</strong>, y navega hasta la carpeta raíz <code className="font-bold font-mono">public_html</code>. Copia todo el contenido de la carpeta compilada <code className="font-bold font-mono text-indigo-600">dist/</code> (no arrastres la carpeta base entero, sino los archivos que están adentro) directamente allí. ¡Tu web estará lista!</p>
                    </div>
                  </div>
                </div>
              </div>

              {/* htaccess e información técnica adicional */}
              <div>
                <h3 className="text-xs font-bold uppercase tracking-wider text-slate-800 flex items-center gap-1.5 mb-2">
                  <FileJson size={14} className="text-slate-600" />
                  Estructura pre-instalada de Redirección (.htaccess)
                </h3>
                <div className="bg-slate-900 rounded-xl p-3 text-[10px] font-mono text-slate-300 relative select-all leading-relaxed overflow-x-auto">
                  <pre>{`<IfModule mod_rewrite.c>
  RewriteEngine On
  RewriteBase /
  RewriteRule ^index\\.html$ - [L]
  RewriteCond %{REQUEST_FILENAME} !-f
  RewriteCond %{REQUEST_FILENAME} !-d
  RewriteRule . /index.html [L]
</IfModule>`}</pre>
                </div>
              </div>

            </div>

            {/* Footer de Acciones del Modal */}
            <div className="bg-slate-50 p-4 border-t border-slate-100 flex flex-wrap items-center justify-between gap-3 shrink-0">
              <button
                onClick={() => {
                  if (confirm("¿Estás seguro de que deseas vaciar los cursos de este navegador? Se restablecerán por completo a la plantilla de diseño original.")) {
                    localStorage.removeItem('fundes_courses');
                    localStorage.removeItem('fundes_hostinger_local_mode');
                    window.location.reload();
                  }
                }}
                className="px-3 py-2 text-[10px] font-bold text-rose-600 bg-rose-50 border border-rose-100 hover:bg-rose-100 rounded-xl transition-all cursor-pointer"
              >
                Limpiar datos locales del navegador
              </button>
              <button
                onClick={() => setIsHostingerModalOpen(false)}
                className="px-5 py-2 text-xs font-bold text-white bg-slate-800 hover:bg-slate-900 rounded-xl shadow-lg shadow-slate-900/10 transition-all cursor-pointer"
              >
                Cerrar Ayuda Despliegue
              </button>
            </div>

          </div>
        </div>
      )}

    </div>
  );
}
