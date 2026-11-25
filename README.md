```html
<!DOCTYPE html>
<html lang="es" class="dark">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
  <title>Control de Puerta · Boliche</title>

  <!-- Google Fonts: Inter -->
  <link rel="preconnect" href="https://fonts.googleapis.com" />
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap" rel="stylesheet" />

  <!-- Tailwind CSS CDN -->
  <script src="https://cdn.tailwindcss.com"></script>
  <script>
    tailwind.config = {
      darkMode: 'class',
      theme: {
        extend: {
          fontFamily: {
            sans: ['Inter', 'system-ui', '-apple-system', 'sans-serif'],
          },
          colors: {
            amoled: '#09090b',
            panel: '#111318',
            card: '#161922',
          }
        }
      }
    };
  </script>

  <!-- Chart.js CDN -->
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.3/dist/chart.umd.min.js"></script>

  <!-- SheetJS (xlsx) CDN para exportación real a Excel -->
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>

  <style>
    /* Prevención de zoom por doble toque y selección accidental en móviles */
    html, body {
      background-color: #09090b;
      color: #f4f4f5;
      touch-action: manipulation;
      -webkit-tap-highlight-color: transparent;
      overscroll-behavior-y: none;
    }

    .no-select {
      -webkit-user-select: none;
      user-select: none;
      touch-action: manipulation;
    }

    .tabular-nums {
      font-variant-numeric: tabular-nums;
    }

    /* Animación sutil de pulso al registrar toque */
    @keyframes tapPop {
      0% { transform: scale(1); }
      40% { transform: scale(0.965); }
      100% { transform: scale(1); }
    }
    .tap-pop {
      animation: tapPop 140ms ease-out;
    }

    /* Scrollbar oscuro discreto */
    ::-webkit-scrollbar {
      width: 6px;
      height: 6px;
    }
    ::-webkit-scrollbar-track {
      background: #09090b;
    }
    ::-webkit-scrollbar-thumb {
      background: #27272a;
      border-radius: 999px;
    }
  </style>
</head>
<body class="bg-amoled text-zinc-100 font-sans min-h-screen flex flex-col antialiased selection:bg-sky-500/30">

  <!-- ==================== BARRA SUPERIOR DE NAVEGACIÓN ==================== -->
  <header class="sticky top-0 z-30 bg-amoled/95 backdrop-blur-md border-b border-zinc-800/80 px-3 py-2.5">
    <div class="max-w-lg mx-auto flex items-center justify-between gap-2">
      <!-- Estado de sincronización y código de sala -->
      <button
        id="btnOpenSync"
        type="button"
        class="no-select flex items-center gap-2 bg-zinc-900 hover:bg-zinc-800 active:scale-95 transition border border-zinc-800 rounded-xl px-2.5 py-1.5 text-left"
        title="Sincronizar dispositivos"
      >
        <span id="syncDot" class="w-2.5 h-2.5 rounded-full bg-amber-400 animate-pulse shrink-0"></span>
        <div class="leading-tight">
          <div class="flex items-center gap-1.5">
            <span class="text-[10px] uppercase tracking-wider text-zinc-400 font-semibold">Sala</span>
            <span id="headerRoomCode" class="text-xs font-extrabold tracking-widest text-zinc-100 tabular-nums">------</span>
          </div>
          <span id="syncStatusText" class="text-[10px] text-zinc-400 block">Conectando...</span>
        </div>
      </button>

      <!-- Reloj + acciones rápidas -->
      <div class="flex items-center gap-1.5">
        <!-- Reloj actual y cierre configurado -->
        <div class="hidden xs:flex flex-col items-end mr-1 leading-tight">
          <span id="liveClock" class="text-xs font-bold text-zinc-200 tabular-nums">--:--</span>
          <span id="closingTimeBadge" class="text-[10px] text-zinc-500 tabular-nums">Cierre 05:00</span>
        </div>

        <!-- Botón Gráficos / Reporte -->
        <button
          id="btnOpenReport"
          type="button"
          class="no-select flex items-center gap-1.5 bg-indigo-600/20 hover:bg-indigo-600/30 active:scale-95 transition text-indigo-300 border border-indigo-500/40 rounded-xl px-3 py-2 text-xs font-bold"
        >
          <svg class="w-4 h-4" fill="none" stroke="currentColor" stroke-width="2.2" viewBox="0 0 24 24">
            <path stroke-linecap="round" stroke-linejoin="round" d="M3 13.125C3 12.504 3.504 12 4.125 12h2.25c.621 0 1.125.504 1.125 1.125v6.75C7.5 20.496 6.996 21 6.375 21h-2.25A1.125 1.125 0 013 19.875v-6.75zM9.75 8.625c0-.621.504-1.125 1.125-1.125h2.25c.621 0 1.125.504 1.125 1.125v11.25c0 .621-.504 1.125-1.125 1.125h-2.25a1.125 1.125 0 01-1.125-1.125V8.625zM16.5 4.125c0-.621.504-1.125 1.125-1.125h2.25C20.496 3 21 3.504 21 4.125v15.75c0 .621-.504 1.125-1.125 1.125h-2.25a1.125 1.125 0 01-1.125-1.125V4.125z" />
          </svg>
          <span>Reporte</span>
        </button>

        <!-- Botón Configuración -->
        <button
          id="btnOpenSettings"
          type="button"
          aria-label="Configuración"
          class="no-select p-2 bg-zinc-900 hover:bg-zinc-800 active:scale-95 transition text-zinc-300 border border-zinc-800 rounded-xl"
        >
          <svg class="w-5 h-5" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
            <path stroke-linecap="round" stroke-linejoin="round" d="M10.325 4.317c.426-1.756 2.924-1.756 3.35 0a1.724 1.724 0 002.573 1.066c1.543-.94 3.31.826 2.37 2.37a1.724 1.724 0 001.065 2.572c1.756.426 1.756 2.924 0 3.35a1.724 1.724 0 00-1.066 2.573c.94 1.543-.826 3.31-2.37 2.37a1.724 1.724 0 00-2.572 1.065c-.426 1.756-2.924 1.756-3.35 0a1.724 1.724 0 00-2.573-1.066c-1.543.94-3.31-.826-2.37-2.37a1.724 1.724 0 00-1.065-2.572c-1.756-.426-1.756-2.924 0-3.35a1.724 1.724 0 001.066-2.573c-.94-1.543.826-3.31 2.37-2.37.996.608 2.296.07 2.572-1.065z" />
            <path stroke-linecap="round" stroke-linejoin="round" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z" />
          </svg>
        </button>
      </div>
    </div>
  </header>

  <!-- ==================== BANNER DE BLOQUEO POR CIERRE AUTOMÁTICO ==================== -->
  <div id="lockedBanner" class="hidden bg-amber-500/15 border-b border-amber-500/40 px-3 py-2">
    <div class="max-w-lg mx-auto flex items-center justify-between gap-2">
      <div class="flex items-center gap-2 text-amber-200 text-xs font-semibold">
        <svg class="w-4 h-4 text-amber-400 shrink-0" fill="none" stroke="currentColor" stroke-width="2.2" viewBox="0 0 24 24">
          <path stroke-linecap="round" stroke-linejoin="round" d="M16.5 10.5V6.75a4.5 4.5 0 10-9 0v3.75m-.75 11.25h10.5a2.25 2.25 0 002.25-2.25v-6.75a2.25 2.25 0 00-2.25-2.25H6.75a2.25 2.25 0 00-2.25 2.25v6.75a2.25 2.25 0 002.25 2.25z" />
        </svg>
        <span id="lockedBannerText">Noche cerrada · Contador bloqueado</span>
      </div>
      <button
        id="btnUnlockCounter"
        type="button"
        class="no-select bg-amber-500 hover:bg-amber-400 active:scale-95 transition text-zinc-950 font-bold text-xs px-3 py-1.5 rounded-lg shrink-0"
      >
        Desbloquear
      </button>
    </div>
  </div>

  <!-- ==================== VISTA PRINCIPAL: CONTADOR ==================== -->
  <main id="counterView" class="flex-1 flex flex-col max-w-lg w-full mx-auto px-3 pt-3 pb-5 gap-3">

    <!-- PANEL SUPERIOR SIEMPRE VISIBLE (TOTALES + PROPORCIÓN) -->
    <section class="bg-panel border border-zinc-800/90 rounded-2xl p-3.5 shadow-xl shadow-black/40">
      <!-- Fila principal de 3 métricas -->
      <div class="grid grid-cols-3 gap-2 items-center text-center">
        <!-- Total Ingresos -->
        <div class="bg-zinc-900/90 border border-zinc-800/80 rounded-xl py-2.5 px-2">
          <p class="text-[10px] uppercase tracking-wider font-bold text-emerald-400/90">Ingresos</p>
          <p id="statTotalIn" class="text-2xl sm:text-3xl font-extrabold text-zinc-100 tabular-nums leading-tight mt-0.5">0</p>
        </div>

        <!-- Personas Adentro (Protagonista) -->
        <div class="bg-gradient-to-b from-zinc-800/90 to-zinc-900 border border-zinc-600/60 rounded-xl py-2.5 px-2 shadow-inner">
          <p class="text-[10px] uppercase tracking-wider font-extrabold text-white">Adentro</p>
          <p id="statInsideTotal" class="text-3xl sm:text-4xl font-black text-white tabular-nums leading-tight mt-0.5">0</p>
        </div>

        <!-- Total Salidas -->
        <div class="bg-zinc-900/90 border border-zinc-800/80 rounded-xl py-2.5 px-2">
          <p class="text-[10px] uppercase tracking-wider font-bold text-amber-400/90">Salidas</p>
          <p id="statTotalOut" class="text-2xl sm:text-3xl font-extrabold text-zinc-100 tabular-nums leading-tight mt-0.5">0</p>
        </div>
      </div>

      <!-- Desglose Adentro Hombres vs Mujeres + Barra de Proporción -->
      <div class="mt-3 pt-3 border-t border-zinc-800/80">
        <div class="flex items-center justify-between text-xs font-bold mb-1.5">
          <div class="flex items-center gap-1.5 text-sky-400">
            <span class="w-2 h-2 rounded-full bg-sky-400"></span>
            <span>Hombres:</span>
            <span id="statInsideH" class="text-sm font-extrabold text-sky-300 tabular-nums">0</span>
            <span id="statPctH" class="text-[11px] font-semibold text-sky-400/80 tabular-nums">(0%)</span>
          </div>
          <div class="flex items-center gap-1.5 text-pink-400">
            <span id="statPctM" class="text-[11px] font-semibold text-pink-400/80 tabular-nums">(0%)</span>
            <span id="statInsideM" class="text-sm font-extrabold text-pink-300 tabular-nums">0</span>
            <span>:Mujeres</span>
            <span class="w-2 h-2 rounded-full bg-pink-500"></span>
          </div>
        </div>

        <!-- Barra visual de proporción -->
        <div class="w-full h-2.5 bg-zinc-800 rounded-full overflow-hidden flex">
          <div id="barH" class="h-full bg-sky-400 transition-all duration-300" style="width: 50%"></div>
          <div id="barM" class="h-full bg-pink-500 transition-all duration-300" style="width: 50%"></div>
        </div>

        <!-- Reloj visible en pantallas muy angostas -->
        <div class="flex items-center justify-between mt-2 text-[11px] text-zinc-500">
          <span id="mobileClockRow" class="tabular-nums">Hora: --:-- · Cierre auto: 05:00</span>
          <span id="lastActionText" class="truncate max-w-[55%] text-right text-zinc-400">Listo para contar</span>
        </div>
      </div>
    </section>

    <!-- ==================== SECCIÓN 1: INGRESOS (BOTONES GRANDES) ==================== -->
    <section class="flex-1 flex flex-col gap-2">
      <div class="flex items-center justify-between px-1">
        <div class="flex items-center gap-1.5">
          <span class="w-2 h-2 rounded-full bg-emerald-400"></span>
          <h2 class="text-xs font-extrabold uppercase tracking-widest text-emerald-400">Ingresos (+1)</h2>
        </div>
        <span class="text-[11px] text-zinc-500">Tocá el bloque para sumar</span>
      </div>

      <div class="grid grid-cols-2 gap-3 flex-1 min-h-[215px]">
        <!-- BLOQUE INGRESO HOMBRES -->
        <div class="flex flex-col rounded-2xl border-2 border-sky-500/60 bg-gradient-to-b from-sky-500/20 via-sky-950/50 to-zinc-950 overflow-hidden shadow-lg shadow-sky-950/30">
          <button
            id="btnInH"
            type="button"
            data-action="in-H"
            class="counter-btn no-select flex-1 flex flex-col items-center justify-center p-4 active:bg-sky-500/30 transition cursor-pointer focus:outline-none"
          >
            <!-- Ícono Hombre -->
            <div class="w-11 h-11 rounded-full bg-sky-500/20 border border-sky-400/40 flex items-center justify-center mb-2 text-sky-300">
              <svg class="w-6 h-6" fill="none" stroke="currentColor" stroke-width="2.2" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" d="M15.75 6a3.75 3.75 0 11-7.5 0 3.75 3.75 0 017.5 0zM4.501 20.118a7.5 7.5 0 0114.998 0A17.933 17.933 0 0112 21.75c-2.676 0-5.216-.584-7.499-1.632z" />
              </svg>
            </div>
            <span class="text-sm font-black uppercase tracking-wider text-sky-300">Hombres</span>
            <span class="text-4xl sm:text-5xl font-black text-white mt-1 leading-none">+1</span>
            <span class="mt-3 text-xs font-semibold text-sky-200/80 bg-sky-950/80 border border-sky-500/30 px-2.5 py-0.5 rounded-full tabular-nums">
              Ingresos: <strong id="countInH" class="text-white">0</strong>
            </span>
          </button>

          <!-- Zona separada de corrección -1 -->
          <div class="bg-zinc-950/90 border-t border-sky-500/25 p-1.5 flex items-center justify-between px-2.5">
            <span class="text-[10px] text-zinc-500 font-medium">Corregir</span>
            <button
              id="btnUndoInH"
              type="button"
              class="no-select bg-zinc-900 hover:bg-zinc-800 active:scale-90 transition border border-zinc-700 text-sky-300 font-extrabold text-xs px-3 py-1.5 rounded-lg tabular-nums"
            >
              -1
            </button>
          </div>
        </div>

        <!-- BLOQUE INGRESO MUJERES -->
        <div class="flex flex-col rounded-2xl border-2 border-pink-500/60 bg-gradient-to-b from-pink-500/20 via-fuchsia-950/50 to-zinc-950 overflow-hidden shadow-lg shadow-pink-950/30">
          <button
            id="btnInM"
            type="button"
            data-action="in-M"
            class="counter-btn no-select flex-1 flex flex-col items-center justify-center p-4 active:bg-pink-500/30 transition cursor-pointer focus:outline-none"
          >
            <!-- Ícono Mujer -->
            <div class="w-11 h-11 rounded-full bg-pink-500/20 border border-pink-400/40 flex items-center justify-center mb-2 text-pink-300">
              <svg class="w-6 h-6" fill="none" stroke="currentColor" stroke-width="2.2" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" d="M15.75 6a3.75 3.75 0 11-7.5 0 3.75 3.75 0 017.5 0zM4.501 20.118a7.5 7.5 0 0114.998 0A17.933 17.933 0 0112 21.75c-2.676 0-5.216-.584-7.499-1.632z" />
              </svg>
            </div>
            <span class="text-sm font-black uppercase tracking-wider text-pink-300">Mujeres</span>
            <span class="text-4xl sm:text-5xl font-black text-white mt-1 leading-none">+1</span>
            <span class="mt-3 text-xs font-semibold text-pink-200/80 bg-pink-950/80 border border-pink-500/30 px-2.5 py-0.5 rounded-full tabular-nums">
              Ingresos: <strong id="countInM" class="text-white">0</strong>
            </span>
          </button>

          <!-- Zona separada de corrección -1 -->
          <div class="bg-zinc-950/90 border-t border-pink-500/25 p-1.5 flex items-center justify-between px-2.5">
            <span class="text-[10px] text-zinc-500 font-medium">Corregir</span>
            <button
              id="btnUndoInM"
              type="button"
              class="no-select bg-zinc-900 hover:bg-zinc-800 active:scale-90 transition border border-zinc-700 text-pink-300 font-extrabold text-xs px-3 py-1.5 rounded-lg tabular-nums"
            >
              -1
            </button>
          </div>
        </div>
      </div>
    </section>

    <!-- ==================== SECCIÓN 2: SALIDAS (CLARAMENTE DIFERENCIADA) ==================== -->
    <section class="flex flex-col gap-2 pt-1">
      <div class="flex items-center justify-between px-1">
        <div class="flex items-center gap-1.5">
          <span class="w-2 h-2 rounded-full bg-amber-400"></span>
          <h2 class="text-xs font-extrabold uppercase tracking-widest text-amber-400">Salidas</h2>
        </div>
        <span class="text-[11px] text-zinc-500">Registra egreso del local</span>
      </div>

      <div class="grid grid-cols-2 gap-3">
        <!-- BLOQUE SALIDA HOMBRES -->
        <div class="flex flex-col rounded-2xl border border-zinc-700/90 bg-panel overflow-hidden">
          <button
            id="btnOutH"
            type="button"
            data-action="out-H"
            class="counter-btn no-select flex flex-col items-center justify-center py-4 px-3 active:bg-zinc-800 transition cursor-pointer focus:outline-none"
          >
            <div class="flex items-center gap-1.5 text-sky-400">
              <svg class="w-4 h-4 text-amber-400" fill="none" stroke="currentColor" stroke-width="2.4" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" d="M15.75 9V5.25A2.25 2.25 0 0013.5 3h-6a2.25 2.25 0 00-2.25 2.25v13.5A2.25 2.25 0 007.5 21h6a2.25 2.25 0 002.25-2.25V15m3 0l3-3m0 0l-3-3m3 3H9" />
              </svg>
              <span class="text-xs font-extrabold uppercase tracking-wider">Salida Hombres</span>
            </div>
            <span class="text-2xl font-black text-zinc-100 mt-1.5 leading-none">+1 <span class="text-xs font-semibold text-amber-400">salida</span></span>
            <span class="mt-2 text-[11px] font-semibold text-zinc-400 bg-zinc-900 border border-zinc-800 px-2.5 py-0.5 rounded-full tabular-nums">
              Salidas: <strong id="countOutH" class="text-zinc-200">0</strong>
            </span>
          </button>

          <div class="bg-zinc-950/80 border-t border-zinc-800 p-1.5 flex items-center justify-between px-2.5">
            <span class="text-[10px] text-zinc-500 font-medium">Corregir</span>
            <button
              id="btnUndoOutH"
              type="button"
              class="no-select bg-zinc-900 hover:bg-zinc-800 active:scale-90 transition border border-zinc-700 text-amber-300 font-extrabold text-xs px-3 py-1.5 rounded-lg tabular-nums"
            >
              -1
            </button>
          </div>
        </div>

        <!-- BLOQUE SALIDA MUJERES -->
        <div class="flex flex-col rounded-2xl border border-zinc-700/90 bg-panel overflow-hidden">
          <button
            id="btnOutM"
            type="button"
            data-action="out-M"
            class="counter-btn no-select flex flex-col items-center justify-center py-4 px-3 active:bg-zinc-800 transition cursor-pointer focus:outline-none"
          >
            <div class="flex items-center gap-1.5 text-pink-400">
              <svg class="w-4 h-4 text-amber-400" fill="none" stroke="currentColor" stroke-width="2.4" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" d="M15.75 9V5.25A2.25 2.25 0 0013.5 3h-6a2.25 2.25 0 00-2.25 2.25v13.5A2.25 2.25 0 007.5 21h6a2.25 2.25 0 002.25-2.25V15m3 0l3-3m0 0l-3-3m3 3H9" />
              </svg>
              <span class="text-xs font-extrabold uppercase tracking-wider">Salida Mujeres</span>
            </div>
            <span class="text-2xl font-black text-zinc-100 mt-1.5 leading-none">+1 <span class="text-xs font-semibold text-amber-400">salida</span></span>
            <span class="mt-2 text-[11px] font-semibold text-zinc-400 bg-zinc-900 border border-zinc-800 px-2.5 py-0.5 rounded-full tabular-nums">
              Salidas: <strong id="countOutM" class="text-zinc-200">0</strong>
            </span>
          </button>

          <div class="bg-zinc-950/80 border-t border-zinc-800 p-1.5 flex items-center justify-between px-2.5">
            <span class="text-[10px] text-zinc-500 font-medium">Corregir</span>
            <button
              id="btnUndoOutM"
              type="button"
              class="no-select bg-zinc-900 hover:bg-zinc-800 active:scale-90 transition border border-zinc-700 text-amber-300 font-extrabold text-xs px-3 py-1.5 rounded-lg tabular-nums"
            >
              -1
            </button>
          </div>
        </div>
      </div>
    </section>
  </main>

  <!-- ==================== VISTA DE REPORTE Y GRÁFICOS ==================== -->
  <section id="reportView" class="hidden flex-1 flex-col max-w-2xl w-full mx-auto px-3 pt-3 pb-10 gap-4">
    <!-- Barra superior del reporte -->
    <div class="flex items-center justify-between gap-2">
      <button
        id="btnBackToCounter"
        type="button"
        class="no-select flex items-center gap-1.5 bg-zinc-900 hover:bg-zinc-800 active:scale-95 transition border border-zinc-800 rounded-xl px-3 py-2 text-xs font-bold text-zinc-200"
      >
        <svg class="w-4 h-4" fill="none" stroke="currentColor" stroke-width="2.4" viewBox="0 0 24 24">
          <path stroke-linecap="round" stroke-linejoin="round" d="M15.75 19.5L8.25 12l7.5-7.5" />
        </svg>
        <span>Volver al contador</span>
      </button>

      <div class="text-right">
        <h1 class="text-sm font-extrabold text-white">Reporte de la Noche</h1>
        <p id="reportMetaSubtitle" class="text-[11px] text-zinc-400">Sala ------ · Intervalo 15 min</p>
      </div>
    </div>

    <!-- Botones de Exportación (WhatsApp PNG + Descargar PNG + Excel) -->
    <div class="grid grid-cols-1 sm:grid-cols-3 gap-2">
      <button
        id="btnShareWhatsApp"
        type="button"
        class="no-select flex items-center justify-center gap-2 bg-emerald-600 hover:bg-emerald-500 active:scale-95 transition text-white font-bold text-xs py-3 px-4 rounded-xl shadow-lg shadow-emerald-950/40"
      >
        <svg class="w-4 h-4" fill="none" stroke="currentColor" stroke-width="2.2" viewBox="0 0 24 24">
          <path stroke-linecap="round" stroke-linejoin="round" d="M7.217 10.907a2.25 2.25 0 100 2.186m0-2.186c.18.324.283.696.283 1.093s-.103.77-.283 1.093m0-2.186l9.566-5.314m-9.566 7.5l9.566 5.314m0 0a2.25 2.25 0 103.935 2.186 2.25 2.25 0 00-3.935-2.186zm0-12.814a2.25 2.25 0 103.933-2.185 2.25 2.25 0 00-3.933 2.185z" />
        </svg>
        <span>Enviar PNG a WhatsApp</span>
      </button>

      <button
        id="btnDownloadPng"
        type="button"
        class="no-select flex items-center justify-center gap-2 bg-zinc-800 hover:bg-zinc-700 active:scale-95 transition text-zinc-100 border border-zinc-700 font-bold text-xs py-3 px-4 rounded-xl"
      >
        <svg class="w-4 h-4" fill="none" stroke="currentColor" stroke-width="2.2" viewBox="0 0 24 24">
          <path stroke-linecap="round" stroke-linejoin="round" d="M3 16.5v2.25A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75V16.5M16.5 12L12 16.5m0 0L7.5 12m4.5 4.5V3" />
        </svg>
        <span>Descargar Imagen PNG</span>
      </button>

      <button
        id="btnExportExcel"
        type="button"
        class="no-select flex items-center justify-center gap-2 bg-sky-600/20 hover:bg-sky-600/30 active:scale-95 transition text-sky-300 border border-sky-500/40 font-bold text-xs py-3 px-4 rounded-xl"
      >
        <svg class="w-4 h-4" fill="none" stroke="currentColor" stroke-width="2.2" viewBox="0 0 24 24">
          <path stroke-linecap="round" stroke-linejoin="round" d="M19.5 14.25v-2.625a3.375 3.375 0 00-3.375-3.375h-1.5A1.125 1.125 0 0113.5 7.125v-1.5a3.375 3.375 0 00-3.375-3.375H8.25m.75 12l3 3m0 0l3-3m-3 3v-6m-1.5-9H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 00-9-9z" />
        </svg>
        <span>Exportar Excel (.xlsx)</span>
      </button>
    </div>

    <!-- Tarjeta de Métricas Clave de la Noche -->
    <div class="grid grid-cols-2 gap-2.5">
      <div class="bg-panel border border-zinc-800 rounded-2xl p-3">
        <span class="text-[10px] uppercase tracking-wider font-bold text-zinc-400">Pico Máximo Adentro</span>
        <p id="kpiPeakInside" class="text-2xl font-black text-white tabular-nums mt-0.5">0</p>
        <p id="kpiPeakInsideTime" class="text-[11px] text-emerald-400 font-semibold mt-0.5">Horario: --:--</p>
      </div>

      <div class="bg-panel border border-zinc-800 rounded-2xl p-3">
        <span class="text-[10px] uppercase tracking-wider font-bold text-zinc-400">Total Ingresos</span>
        <p id="kpiTotalEntries" class="text-2xl font-black text-white tabular-nums mt-0.5">0</p>
        <p id="kpiTotalSplit" class="text-[11px] font-semibold mt-0.5 tabular-nums">
          <span class="text-sky-400">0 H (0%)</span> · <span class="text-pink-400">0 M (0%)</span>
        </p>
      </div>

      <div class="bg-panel border border-zinc-800 rounded-2xl p-3">
        <span class="text-[10px] uppercase tracking-wider font-bold text-pink-400/90">Pico Ingreso Mujeres</span>
        <p id="kpiPeakInM" class="text-xl font-black text-pink-300 tabular-nums mt-0.5">0 / franja</p>
        <p id="kpiPeakInMTime" class="text-[11px] text-zinc-400 font-medium mt-0.5">Horario: --:--</p>
      </div>

      <div class="bg-panel border border-zinc-800 rounded-2xl p-3">
        <span class="text-[10px] uppercase tracking-wider font-bold text-sky-400/90">Pico Ingreso Hombres</span>
        <p id="kpiPeakInH" class="text-xl font-black text-sky-300 tabular-nums mt-0.5">0 / franja</p>
        <p id="kpiPeakInHTime" class="text-[11px] text-zinc-400 font-medium mt-0.5">Horario: --:--</p>
      </div>
    </div>

    <!-- Estado vacío si no hay datos -->
    <div id="reportEmptyState" class="hidden bg-panel border border-dashed border-zinc-700 rounded-2xl p-6 text-center">
      <p class="text-sm font-bold text-zinc-200">Todavía no hay registros en esta noche</p>
      <p class="text-xs text-zinc-400 mt-1">Empezá a contar ingresos o cargá una noche simulada para probar los gráficos y exportaciones.</p>
      <button
        id="btnEmptyLoadSim"
        type="button"
        class="no-select mt-4 bg-indigo-600 hover:bg-indigo-500 active:scale-95 transition text-white font-bold text-xs px-4 py-2.5 rounded-xl"
      >
        Cargar noche de prueba (simulación)
      </button>
    </div>

    <!-- Contenedor de Gráficos -->
    <div id="chartsContainer" class="flex flex-col gap-4">
      <!-- GRÁFICO 1: Ocupación y Proporción -->
      <div class="bg-panel border border-zinc-800 rounded-2xl p-3.5">
        <div class="flex flex-wrap items-center justify-between gap-2 mb-3">
          <div>
            <h3 class="text-xs font-extrabold uppercase tracking-wider text-zinc-100">1. Ocupación y Proporción</h3>
            <p class="text-[11px] text-zinc-400">Personas adentro en simultáneo a lo largo de la noche</p>
          </div>
          <label class="no-select flex items-center gap-1.5 text-[11px] font-semibold text-zinc-300 bg-zinc-900 border border-zinc-700 px-2.5 py-1 rounded-lg cursor-pointer">
            <input id="toggleShowTotalCurve" type="checkbox" checked class="accent-amber-400 rounded" />
            <span>Mostrar Total</span>
          </label>
        </div>
        <div class="relative w-full h-64 sm:h-72">
          <canvas id="chartOccupancy"></canvas>
        </div>
      </div>

      <!-- GRÁFICO 2: Picos de Ingreso por Franja Horaria -->
      <div class="bg-panel border border-zinc-800 rounded-2xl p-3.5">
        <div class="flex items-center justify-between gap-2 mb-3">
          <div>
            <h3 class="text-xs font-extrabold uppercase tracking-wider text-zinc-100">2. Picos de Ingreso por Franja</h3>
            <p id="chartPeaksSubtitle" class="text-[11px] text-zinc-400">Cantidad de ingresos cada 15 minutos</p>
          </div>
          <!-- Selector rápido de intervalo -->
          <div class="flex bg-zinc-900 border border-zinc-800 rounded-lg p-0.5 text-[11px] font-bold">
            <button id="btnBucket15" type="button" class="no-select px-2 py-1 rounded-md bg-zinc-700 text-white">15m</button>
            <button id="btnBucket30" type="button" class="no-select px-2 py-1 rounded-md text-zinc-400">30m</button>
          </div>
        </div>
        <div class="relative w-full h-64 sm:h-72">
          <canvas id="chartPeaks"></canvas>
        </div>
      </div>

      <!-- GRÁFICO 3: Ingresos Acumulados -->
      <div class="bg-panel border border-zinc-800 rounded-2xl p-3.5">
        <div class="mb-3">
          <h3 class="text-xs font-extrabold uppercase tracking-wider text-zinc-100">3. Ingresos Acumulados</h3>
          <p class="text-[11px] text-zinc-400">Evolución total acumulada de Hombres vs Mujeres</p>
        </div>
        <div class="relative w-full h-60 sm:h-64">
          <canvas id="chartCumulative"></canvas>
        </div>
      </div>
    </div>
  </section>

  <!-- ==================== MODAL: SINCRONIZACIÓN MULTI-DISPOSITIVO ==================== -->
  <div id="syncModal" class="hidden fixed inset-0 z-50 bg-black/80 backdrop-blur-sm flex items-end sm:items-center justify-center p-3">
    <div class="bg-panel border border-zinc-800 rounded-2xl w-full max-w-md p-4 shadow-2xl">
      <div class="flex items-center justify-between mb-3">
        <div class="flex items-center gap-2">
          <span id="modalSyncDot" class="w-2.5 h-2.5 rounded-full bg-emerald-400"></span>
          <h3 class="text-base font-extrabold text-white">Sincronizar Celulares</h3>
        </div>
        <button id="btnCloseSync" type="button" class="no-select p-1.5 text-zinc-400 hover:text-white rounded-lg">
          <svg class="w-5 h-5" fill="none" stroke="currentColor" stroke-width="2.2" viewBox="0 0 24 24">
            <path stroke-linecap="round" stroke-linejoin="round" d="M6 18L18 6M6 6l12 12" />
          </svg>
        </button>
      </div>

      <p class="text-xs text-zinc-400 mb-3">
        Compartí este código o enlace con otro operador en otra puerta para que ambos celulares cuenten sobre la misma noche en tiempo real sin pisarse.
      </p>

      <!-- Código actual -->
      <div class="bg-zinc-950 border border-zinc-800 rounded-xl p-3 text-center mb-3">
        <span class="text-[10px] uppercase tracking-widest text-zinc-500 font-bold">Código de Noche Activo</span>
        <div id="modalRoomCodeDisplay" class="text-3xl font-black tracking-[0.25em] text-sky-400 mt-1 tabular-nums">------</div>
        <div class="grid grid-cols-2 gap-2 mt-3">
          <button id="btnCopyCode" type="button" class="no-select bg-zinc-800 hover:bg-zinc-700 active:scale-95 transition text-zinc-100 text-xs font-bold py-2 rounded-lg">
            Copiar Código
          </button>
          <button id="btnCopyLink" type="button" class="no-select bg-sky-600 hover:bg-sky-500 active:scale-95 transition text-white text-xs font-bold py-2 rounded-lg">
            Copiar Enlace
          </button>
        </div>
      </div>

      <!-- Unirse a otra sala -->
      <div class="border-t border-zinc-800 pt-3 mb-3">
        <label for="inputJoinCode" class="block text-xs font-bold text-zinc-300 mb-1.5">Unirse a una sala existente</label>
        <div class="flex gap-2">
          <input
            id="inputJoinCode"
            type="text"
            maxlength="8"
            placeholder="Ej: N8K4P2"
            class="flex-1 bg-zinc-950 border border-zinc-700 focus:border-sky-500 focus:outline-none rounded-xl px-3 py-2 text-sm font-bold uppercase tracking-widest text-white"
          />
          <button id="btnJoinRoom" type="button" class="no-select bg-indigo-600 hover:bg-indigo-500 active:scale-95 transition text-white text-xs font-bold px-4 py-2 rounded-xl">
            Unirme
          </button>
        </div>
      </div>

      <!-- Crear nueva sala + info de operador -->
      <div class="flex items-center justify-between gap-2 pt-2 border-t border-zinc-800">
        <button id="btnCreateNewRoom" type="button" class="no-select text-xs font-bold text-zinc-300 hover:text-white underline underline-offset-2">
          + Crear sala nueva
        </button>
        <span id="modalConnectionBadge" class="text-[11px] font-semibold text-emerald-400">En línea</span>
      </div>

      <!-- ID completo del dispositivo/operador -->
      <div class="mt-3 pt-2 border-t border-zinc-800/70">
        <p class="text-[10px] text-zinc-500 break-all">
          ID Operador: <span id="modalUserId" class="font-mono text-zinc-400">Inicializando...</span>
        </p>
      </div>
    </div>
  </div>

  <!-- ==================== MODAL: CONFIGURACIÓN ==================== -->
  <div id="settingsModal" class="hidden fixed inset-0 z-50 bg-black/80 backdrop-blur-sm flex items-end sm:items-center justify-center p-3">
    <div class="bg-panel border border-zinc-800 rounded-2xl w-full max-w-md p-4 shadow-2xl">
      <div class="flex items-center justify-between mb-4">
        <h3 class="text-base font-extrabold text-white">Configuración de la Noche</h3>
        <button id="btnCloseSettings" type="button" class="no-select p-1.5 text-zinc-400 hover:text-white rounded-lg">
          <svg class="w-5 h-5" fill="none" stroke="currentColor" stroke-width="2.2" viewBox="0 0 24 24">
            <path stroke-linecap="round" stroke-linejoin="round" d="M6 18L18 6M6 6l12 12" />
          </svg>
        </button>
      </div>

      <div class="space-y-4">
        <!-- Hora de cierre automático -->
        <div class="flex items-center justify-between gap-3 bg-zinc-950 border border-zinc-800 rounded-xl p-3">
          <div>
            <label for="inputClosingTime" class="block text-xs font-bold text-zinc-100">Hora de cierre automático</label>
            <p class="text-[11px] text-zinc-400">Bloquea botones y abre el reporte</p>
          </div>
          <input
            id="inputClosingTime"
            type="time"
            value="05:00"
            class="bg-zinc-900 border border-zinc-700 text-white font-bold text-sm rounded-lg px-2.5 py-1.5 focus:outline-none focus:border-sky-500"
          />
        </div>

        <!-- Intervalo de agrupación de gráficos -->
        <div class="flex items-center justify-between gap-3 bg-zinc-950 border border-zinc-800 rounded-xl p-3">
          <div>
            <label for="selectBucketMinutes" class="block text-xs font-bold text-zinc-100">Franja para gráficos</label>
            <p class="text-[11px] text-zinc-400">Agrupación del eje temporal</p>
          </div>
          <select
            id="selectBucketMinutes"
            class="bg-zinc-900 border border-zinc-700 text-white font-bold text-xs rounded-lg px-3 py-2 focus:outline-none focus:border-sky-500"
          >
            <option value="15">Cada 15 min</option>
            <option value="30">Cada 30 min</option>
          </select>
        </div>

        <!-- Cargar noche de prueba -->
        <div class="bg-zinc-950 border border-indigo-500/30 rounded-xl p-3">
          <p class="text-xs font-bold text-indigo-300">Simulación rápida</p>
          <p class="text-[11px] text-zinc-400 mt-0.5 mb-2.5">
            Genera una noche completa de prueba (23:30 a 05:00) con picos realistas de mujeres y hombres para probar los gráficos y exportaciones ahora mismo.
          </p>
          <button
            id="btnLoadSimulation"
            type="button"
            class="no-select w-full bg-indigo-600 hover:bg-indigo-500 active:scale-95 transition text-white font-bold text-xs py-2.5 rounded-xl"
          >
            Cargar noche de prueba (simulación)
          </button>
        </div>

        <!-- Reiniciar noche con doble confirmación -->
        <div class="bg-zinc-950 border border-rose-500/30 rounded-xl p-3">
          <p class="text-xs font-bold text-rose-400">Zona de reinicio</p>
          <p class="text-[11px] text-zinc-400 mt-0.5 mb-2.5">
            Pone todos los contadores en cero para arrancar una noche nueva en esta misma sala.
          </p>

          <!-- Paso 1 -->
          <button
            id="btnResetStep1"
            type="button"
            class="no-select w-full bg-rose-600/20 hover:bg-rose-600/30 active:scale-95 transition border border-rose-500/50 text-rose-300 font-bold text-xs py-2.5 rounded-xl"
          >
            Reiniciar noche
          </button>

          <!-- Paso 2 (doble confirmación) -->
          <div id="resetConfirmBox" class="hidden space-y-2 mt-2 p-2.5 bg-rose-950/50 border border-rose-500/50 rounded-xl">
            <p class="text-[11px] font-bold text-rose-200 text-center">
              ¿Seguro? Esto borra los contadores actuales en todos los celulares sincronizados.
            </p>
            <div class="grid grid-cols-2 gap-2">
              <button id="btnResetCancel" type="button" class="no-select bg-zinc-800 hover:bg-zinc-700 text-zinc-200 font-bold text-xs py-2 rounded-lg">
                Cancelar
              </button>
              <button id="btnResetConfirm" type="button" class="no-select bg-rose-600 hover:bg-rose-500 text-white font-extrabold text-xs py-2 rounded-lg">
                Sí, reiniciar todo
              </button>
            </div>
          </div>
        </div>
      </div>
    </div>
  </div>

  <!-- ==================== TOAST DISCRETO ==================== -->
  <div id="toastContainer" class="fixed bottom-4 left-1/2 -translate-x-1/2 z-[60] pointer-events-none w-[92%] max-w-sm flex flex-col items-center gap-2"></div>

  <!-- ==================== LÓGICA PRINCIPAL + FIREBASE ==================== -->
  <script type="module">
    import { initializeApp } from "https://www.gstatic.com/firebasejs/11.6.1/firebase-app.js";
    import { getAuth, signInAnonymously, signInWithCustomToken } from "https://www.gstatic.com/firebasejs/11.6.1/firebase-auth.js";
    import {
      getFirestore,
      doc,
      getDoc,
      setDoc,
      addDoc,
      deleteDoc,
      onSnapshot,
      collection,
      getDocs
    } from "https://www.gstatic.com/firebasejs/11.6.1/firebase-firestore.js";

    // =========================================================================
    // 1. CONFIGURACIÓN GLOBAL Y ESTADO DE LA APLICACIÓN
    // =========================================================================
    const appId = typeof __app_id !== 'undefined' ? __app_id : 'boliche-door-counter';

    let db = null;
    let auth = null;
    let currentUser = null;
    let firebaseReady = false;

    let unsubRoom = null;
    let unsubEvents = null;

    const state = {
      roomCode: '',
      closingTime: '05:00',
      bucketMinutes: 15,
      isLocked: false,
      resetEpoch: 0,
      lastAutoClosedKey: '',
      events: [], // Lista de eventos { id, tipo: 'in'|'out', genero: 'H'|'M', delta: number, timestamp: number, uid, simulated? }
      activeView: 'counter', // 'counter' | 'report'
      showTotalCurve: true,
    };

    // Instancias de gráficos Chart.js
    const charts = {
      occupancy: null,
      peaks: null,
      cumulative: null,
    };

    // =========================================================================
    // 2. UTILIDADES DE INTERFAZ (TOAST, VIBRACIÓN, PORTAPAPELES, FORMATO)
    // =========================================================================
    function vibrate(ms = 20) {
      try {
        if (navigator.vibrate) navigator.vibrate(ms);
      } catch (_) {}
    }

    function showToast(message, type = 'info') {
      const container = document.getElementById('toastContainer');
      const el = document.createElement('div');

      const colors = {
        info: 'bg-zinc-900/95 border-zinc-700 text-zinc-100',
        warn: 'bg-amber-950/95 border-amber-500/60 text-amber-200',
        success: 'bg-emerald-950/95 border-emerald-500/60 text-emerald-200',
        error: 'bg-rose-950/95 border-rose-500/60 text-rose-200',
      };

      el.className = `pointer-events-auto px-4 py-2.5 rounded-xl border text-xs font-bold shadow-2xl backdrop-blur-md transition-all duration-300 opacity-0 translate-y-2 text-center ${colors[type] || colors.info}`;
      el.textContent = message;
      container.appendChild(el);

      requestAnimationFrame(() => {
        el.classList.remove('opacity-0', 'translate-y-2');
      });

      setTimeout(() => {
        el.classList.add('opacity-0', 'translate-y-2');
        setTimeout(() => el.remove(), 300);
      }, 2600);
    }

    function copyTextFallback(text) {
      const textarea = document.createElement('textarea');
      textarea.value = text;
      textarea.setAttribute('readonly', '');
      textarea.style.position = 'fixed';
      textarea.style.top = '-9999px';
      document.body.appendChild(textarea);
      textarea.select();
      let ok = false;
      try {
        ok = document.execCommand('copy');
      } catch (_) {
        ok = false;
      }
      document.body.removeChild(textarea);
      return ok;
    }

    function generateRoomCode() {
      const chars = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789';
      let code = '';
      for (let i = 0; i < 6; i++) {
        code += chars.charAt(Math.floor(Math.random() * chars.length));
      }
      return code;
    }

    function formatTimeHHMM(ts) {
      const d = new Date(ts);
      const hh = String(d.getHours()).padStart(2, '0');
      const mm = String(d.getMinutes()).padStart(2, '0');
      return `${hh}:${mm}`;
    }

    function formatDateYYYYMMDD(d = new Date()) {
      const y = d.getFullYear();
      const m = String(d.getMonth() + 1).padStart(2, '0');
      const day = String(d.getDate()).padStart(2, '0');
      return `${y}-${m}-${day}`;
    }

    function getRoomFromUrl() {
      try {
        const params = new URLSearchParams(window.location.search);
        const q = params.get('sala');
        if (q && /^[A-Za-z0-9]{4,8}$/.test(q.trim())) return q.trim().toUpperCase();
        if (window.location.hash && window.location.hash.startsWith('#sala=')) {
          const h = window.location.hash.replace('#sala=', '').trim();
          if (/^[A-Za-z0-9]{4,8}$/.test(h)) return h.toUpperCase();
        }
      } catch (_) {}
      return null;
    }

    function updateUrlSafely(roomCode) {
      try {
        const url = new URL(window.location.href);
        url.searchParams.set('sala', roomCode);
        window.history.replaceState({}, '', url.toString());
      } catch (_) {
        // Algunos entornos sandbox restringen replaceState; usar hash como respaldo
        try {
          window.location.hash = `sala=${roomCode}`;
        } catch (__) {}
      }
    }

    function getInviteUrl() {
      try {
        const url = new URL(window.location.href);
        url.searchParams.set('sala', state.roomCode);
        return url.toString();
      } catch (_) {
        return `?sala=${state.roomCode}`;
      }
    }

    // =========================================================================
    // 3. INICIALIZACIÓN DE FIREBASE Y AUTENTICACIÓN ANÓNIMA AUTOMÁTICA
    // =========================================================================
    function setSyncIndicator(status, text) {
      const dot = document.getElementById('syncDot');
      const modalDot = document.getElementById('modalSyncDot');
      const label = document.getElementById('syncStatusText');
      const badge = document.getElementById('modalConnectionBadge');

      const dotStyles = {
        online: 'bg-emerald-400',
        connecting: 'bg-amber-400 animate-pulse',
        local: 'bg-sky-400',
        error: 'bg-rose-500',
      };

      dot.className = `w-2.5 h-2.5 rounded-full shrink-0 ${dotStyles[status] || dotStyles.connecting}`;
      modalDot.className = `w-2.5 h-2.5 rounded-full ${dotStyles[status] || dotStyles.connecting}`;
      label.textContent = text;
      badge.textContent = text;
    }

    async function initFirebase() {
      setSyncIndicator('connecting', 'Conectando...');

      // Verificamos si existe configuración de Firebase provista por el entorno
      if (typeof __firebase_config === 'undefined' || !__firebase_config) {
        firebaseReady = false;
        const urlRoom = getRoomFromUrl() || generateRoomCode();
        state.roomCode = urlRoom;
        updateRoomCodeUI();
        setSyncIndicator('local', 'Modo local');
        document.getElementById('modalUserId').textContent = 'Modo sin nube (local)';
        recalculateAndRender();
        return;
      }

      try {
        const firebaseConfig = typeof __firebase_config === 'string'
          ? JSON.parse(__firebase_config)
          : __firebase_config;

        const app = initializeApp(firebaseConfig);
        db = getFirestore(app);
        auth = getAuth(app);

        // REGLA 3: Autenticar SIEMPRE primero antes de cualquier consulta
        if (typeof __initial_auth_token !== 'undefined' && __initial_auth_token) {
          const cred = await signInWithCustomToken(auth, __initial_auth_token);
          currentUser = cred.user;
        } else {
          const cred = await signInAnonymously(auth);
          currentUser = cred.user;
        }

        if (!currentUser) {
          throw new Error('No se pudo autenticar al operador.');
        }

        firebaseReady = true;
        document.getElementById('modalUserId').textContent = currentUser.uid;

        // Determinar la sala inicial: URL -> preferencia guardada en Firestore -> nueva sala
        let initialCode = getRoomFromUrl();

        if (!initialCode) {
          try {
            const prefRef = doc(db, 'artifacts', appId, 'users', currentUser.uid, 'settings', 'active_session');
            const prefSnap = await getDoc(prefRef);
            if (prefSnap.exists() && prefSnap.data().roomCode) {
              initialCode = String(prefSnap.data().roomCode).toUpperCase();
            }
          } catch (prefErr) {
            console.warn('No se pudo leer preferencia de sala:', prefErr);
          }
        }

        if (!initialCode) {
          initialCode = generateRoomCode();
        }

        await connectToRoom(initialCode);
      } catch (err) {
        console.error('Error iniciando Firebase:', err);
        firebaseReady = false;
        state.roomCode = getRoomFromUrl() || generateRoomCode();
        updateRoomCodeUI();
        setSyncIndicator('local', 'Modo local');
        showToast('Sin conexión a la nube: funcionando en modo local', 'warn');
        recalculateAndRender();
      }
    }

    // =========================================================================
    // 4. CONEXIÓN Y SUSCRIPCIÓN EN TIEMPO REAL A UNA SALA
    // =========================================================================
    function updateRoomCodeUI() {
      document.getElementById('headerRoomCode').textContent = state.roomCode;
      document.getElementById('modalRoomCodeDisplay').textContent = state.roomCode;
      document.getElementById('reportMetaSubtitle').textContent = `Sala ${state.roomCode} · Franjas de ${state.bucketMinutes} min`;
      updateUrlSafely(state.roomCode);
    }

    async function saveUserActiveRoom(roomCode) {
      if (!firebaseReady || !currentUser) return;
      try {
        const prefRef = doc(db, 'artifacts', appId, 'users', currentUser.uid, 'settings', 'active_session');
        await setDoc(prefRef, { roomCode, updatedAt: Date.now() }, { merge: true });
      } catch (err) {
        console.warn('Error guardando sala activa del usuario:', err);
      }
    }

    async function connectToRoom(rawCode) {
      const code = rawCode.trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
      if (!code || code.length < 4) {
        showToast('Ingresá un código válido de sala', 'warn');
        return;
      }

      // Desuscribir listeners anteriores
      if (unsubRoom) { unsubRoom(); unsubRoom = null; }
      if (unsubEvents) { unsubEvents(); unsubEvents = null; }

      state.roomCode = code;
      state.events = [];
      state.isLocked = false;
      state.resetEpoch = 0;
      updateRoomCodeUI();

      if (!firebaseReady || !currentUser) {
        recalculateAndRender();
        return;
      }

      setSyncIndicator('connecting', 'Sincronizando...');
      saveUserActiveRoom(code);

      // REGLA 1: Rutas estrictas bajo /artifacts/{appId}/public/data/{collectionName}
      const roomDocRef = doc(db, 'artifacts', appId, 'public', 'data', 'club_rooms', code);
      const eventsColRef = collection(db, 'artifacts', appId, 'public', 'data', `club_events_${code}`);

      // Verificar si la sala ya existe; si no, crear metadata inicial
      try {
        const snap = await getDoc(roomDocRef);
        if (!snap.exists()) {
          await setDoc(roomDocRef, {
            code,
            closingTime: state.closingTime || '05:00',
            bucketMinutes: state.bucketMinutes || 15,
            isLocked: false,
            resetEpoch: 0,
            lastAutoClosedKey: '',
            createdAt: Date.now(),
          });
        }
      } catch (err) {
        console.error('Error creando/verificando sala:', err);
      }

      // Escuchar cambios de configuración / estado de la sala
      unsubRoom = onSnapshot(
        roomDocRef,
        (docSnap) => {
          if (docSnap.exists()) {
            const data = docSnap.data();
            state.closingTime = data.closingTime || '05:00';
            state.bucketMinutes = Number(data.bucketMinutes) === 30 ? 30 : 15;
            state.isLocked = Boolean(data.isLocked);
            state.resetEpoch = Number(data.resetEpoch) || 0;
            state.lastAutoClosedKey = data.lastAutoClosedKey || '';

            syncSettingsInputsUI();
            recalculateAndRender();
          }
        },
        (error) => {
          console.error('Error en snapshot de sala:', error);
          setSyncIndicator('error', 'Error de sync');
        }
      );

      // Escuchar eventos individuales en tiempo real (REGLA 2: sin orderBy ni where compuesto)
      unsubEvents = onSnapshot(
        eventsColRef,
        (querySnap) => {
          const loaded = [];
          querySnap.forEach((d) => {
            const ev = d.data();
            loaded.push({
              id: d.id,
              tipo: ev.tipo,
              genero: ev.genero,
              delta: Number(ev.delta) || 0,
              timestamp: Number(ev.timestamp) || 0,
              uid: ev.uid || 'op',
              simulated: Boolean(ev.simulated),
            });
          });

          // Ordenamiento en memoria por timestamp ascendente
          loaded.sort((a, b) => a.timestamp - b.timestamp);
          state.events = loaded;
          setSyncIndicator('online', 'En línea · Sincronizado');
          recalculateAndRender();
        },
        (error) => {
          console.error('Error en snapshot de eventos:', error);
          setSyncIndicator('error', 'Sin conexión');
        }
      );
    }

    async function updateRoomMeta(patch) {
      Object.assign(state, patch);
      syncSettingsInputsUI();
      recalculateAndRender();

      if (!firebaseReady || !currentUser) return;
      try {
        const roomDocRef = doc(db, 'artifacts', appId, 'public', 'data', 'club_rooms', state.roomCode);
        await setDoc(roomDocRef, patch, { merge: true });
      } catch (err) {
        console.error('Error actualizando metadata de sala:', err);
      }
    }

    // =========================================================================
    // 5. MOTOR DE CÁLCULO Y VALIDACIÓN DE EVENTOS
    // =========================================================================
    function getActiveEvents() {
      const epoch = state.resetEpoch || 0;
      return state.events
        .filter((ev) => ev.timestamp >= epoch)
        .sort((a, b) => a.timestamp - b.timestamp);
    }

    function computeTotals() {
      const active = getActiveEvents();
      let inH = 0, inM = 0, outH = 0, outM = 0;

      for (const ev of active) {
        if (ev.tipo === 'in' && ev.genero === 'H') inH += ev.delta;
        else if (ev.tipo === 'in' && ev.genero === 'M') inM += ev.delta;
        else if (ev.tipo === 'out' && ev.genero === 'H') outH += ev.delta;
        else if (ev.tipo === 'out' && ev.genero === 'M') outM += ev.delta;
      }

      inH = Math.max(0, inH);
      inM = Math.max(0, inM);
      outH = Math.max(0, outH);
      outM = Math.max(0, outM);

      const insideH = Math.max(0, inH - outH);
      const insideM = Math.max(0, inM - outM);
      const totalIn = inH + inM;
      const totalOut = outH + outM;
      const insideTotal = insideH + insideM;

      const pctH = insideTotal > 0 ? Math.round((insideH / insideTotal) * 100) : 0;
      const pctM = insideTotal > 0 ? 100 - pctH : 0;

      const pctInH = totalIn > 0 ? Math.round((inH / totalIn) * 100) : 0;
      const pctInM = totalIn > 0 ? 100 - pctInH : 0;

      return {
        inH, inM, outH, outM,
        insideH, insideM, insideTotal,
        totalIn, totalOut,
        pctH, pctM,
        pctInH, pctInM,
      };
    }

    async function recordEvent(tipo, genero, delta, btnEl = null) {
      if (state.isLocked) {
        vibrate(40);
        showToast('Noche cerrada: tocá "Desbloquear" para seguir contando', 'warn');
        return;
      }

      const totals = computeTotals();

      // Validaciones lógicas estrictas
      if (delta < 0) {
        if (tipo === 'in' && genero === 'H') {
          if (totals.inH <= 0) { showToast('Ingresos de Hombres ya está en 0', 'warn'); return; }
          if (totals.insideH <= 0) { showToast('No podés restar ingreso: hay 0 hombres adentro', 'warn'); return; }
        }
        if (tipo === 'in' && genero === 'M') {
          if (totals.inM <= 0) { showToast('Ingresos de Mujeres ya está en 0', 'warn'); return; }
          if (totals.insideM <= 0) { showToast('No podés restar ingreso: hay 0 mujeres adentro', 'warn'); return; }
        }
        if (tipo === 'out' && genero === 'H' && totals.outH <= 0) {
          showToast('Salidas de Hombres ya está en 0', 'warn');
          return;
        }
        if (tipo === 'out' && genero === 'M' && totals.outM <= 0) {
          showToast('Salidas de Mujeres ya está en 0', 'warn');
          return;
        }
      } else if (delta > 0 && tipo === 'out') {
        if (genero === 'H' && totals.insideH <= 0) {
          vibrate(35);
          showToast('No hay hombres adentro para registrar salida', 'warn');
          return;
        }
        if (genero === 'M' && totals.insideM <= 0) {
          vibrate(35);
          showToast('No hay mujeres adentro para registrar salida', 'warn');
          return;
        }
      }

      // Feedback táctil y visual inmediato
      vibrate(delta > 0 ? 18 : 28);
      if (btnEl) {
        btnEl.classList.remove('tap-pop');
        void btnEl.offsetWidth;
        btnEl.classList.add('tap-pop');
      }

      const newEvent = {
        tipo,
        genero,
        delta,
        timestamp: Date.now(),
        uid: currentUser ? currentUser.uid : 'local',
      };

      const actionDesc = `${tipo === 'in' ? 'Ingreso' : 'Salida'} ${genero === 'H' ? 'Hombre' : 'Mujer'} (${delta > 0 ? '+1' : '-1'}) · ${formatTimeHHMM(newEvent.timestamp)}`;
      document.getElementById('lastActionText').textContent = actionDesc;

      if (!firebaseReady || !currentUser) {
        state.events.push({ id: `local_${Date.now()}_${Math.random()}`, ...newEvent });
        recalculateAndRender();
        return;
      }

      try {
        const eventsColRef = collection(db, 'artifacts', appId, 'public', 'data', `club_events_${state.roomCode}`);
        await addDoc(eventsColRef, newEvent);
      } catch (err) {
        console.error('Error guardando evento en Firestore:', err);
        showToast('Error al sincronizar el toque', 'error');
      }
    }

    // =========================================================================
    // 6. AGRUPACIÓN TEMPORAL PARA GRÁFICOS Y EXCEL
    // =========================================================================
    function buildTimeSeries() {
      const active = getActiveEvents();
      const bucketMs = state.bucketMinutes * 60 * 1000;

      if (active.length === 0) {
        return {
          buckets: [],
          kpis: {
            peakInside: 0,
            peakInsideTime: '--:--',
            peakInH: 0,
            peakInHTime: '--:--',
            peakInM: 0,
            peakInMTime: '--:--',
          }
        };
      }

      const firstTs = active[0].timestamp;
      const lastTs = active[active.length - 1].timestamp;

      // Evitamos generar miles de franjas vacías si hubo un salto mayor a 18 horas
      const effectiveStart = Math.max(firstTs, lastTs - 18 * 60 * 60 * 1000);
      const startBucket = Math.floor(effectiveStart / bucketMs) * bucketMs;
      let endBucket = Math.floor(lastTs / bucketMs) * bucketMs;
      if (endBucket === startBucket) {
        endBucket = startBucket + bucketMs;
      }

      const map = new Map();
      for (let t = startBucket; t <= endBucket; t += bucketMs) {
        map.set(t, {
          ts: t,
          label: formatTimeHHMM(t),
          inH: 0,
          inM: 0,
          outH: 0,
          outM: 0,
          cumInH: 0,
          cumInM: 0,
          insideH: 0,
          insideM: 0,
          insideTotal: 0,
          pctH: 0,
          pctM: 0,
        });
      }

      // Acumulamos eventos previos a effectiveStart si existieran
      let runInH = 0, runInM = 0, runOutH = 0, runOutM = 0;

      for (const ev of active) {
        const bKey = Math.floor(ev.timestamp / bucketMs) * bucketMs;
        if (bKey < startBucket) {
          if (ev.tipo === 'in' && ev.genero === 'H') runInH += ev.delta;
          if (ev.tipo === 'in' && ev.genero === 'M') runInM += ev.delta;
          if (ev.tipo === 'out' && ev.genero === 'H') runOutH += ev.delta;
          if (ev.tipo === 'out' && ev.genero === 'M') runOutM += ev.delta;
        } else {
          const b = map.get(bKey);
          if (b) {
            if (ev.tipo === 'in' && ev.genero === 'H') b.inH += ev.delta;
            if (ev.tipo === 'in' && ev.genero === 'M') b.inM += ev.delta;
            if (ev.tipo === 'out' && ev.genero === 'H') b.outH += ev.delta;
            if (ev.tipo === 'out' && ev.genero === 'M') b.outM += ev.delta;
          }
        }
      }

      const buckets = [];
      let peakInside = 0, peakInsideTime = '--:--';
      let peakInH = 0, peakInHTime = '--:--';
      let peakInM = 0, peakInMTime = '--:--';

      for (const b of map.values()) {
        runInH = Math.max(0, runInH + b.inH);
        runInM = Math.max(0, runInM + b.inM);
        runOutH = Math.max(0, runOutH + b.outH);
        runOutM = Math.max(0, runOutM + b.outM);

        b.inH = Math.max(0, b.inH);
        b.inM = Math.max(0, b.inM);
        b.outH = Math.max(0, b.outH);
        b.outM = Math.max(0, b.outM);

        b.cumInH = runInH;
        b.cumInM = runInM;
        b.insideH = Math.max(0, runInH - runOutH);
        b.insideM = Math.max(0, runInM - runOutM);
        b.insideTotal = b.insideH + b.insideM;

        b.pctH = b.insideTotal > 0 ? Math.round((b.insideH / b.insideTotal) * 100) : 0;
        b.pctM = b.insideTotal > 0 ? 100 - b.pctH : 0;

        if (b.insideTotal > peakInside) {
          peakInside = b.insideTotal;
          peakInsideTime = b.label;
        }
        if (b.inH > peakInH) {
          peakInH = b.inH;
          peakInHTime = b.label;
        }
        if (b.inM > peakInM) {
          peakInM = b.inM;
          peakInMTime = b.label;
        }

        buckets.push(b);
      }

      return {
        buckets,
        kpis: {
          peakInside,
          peakInsideTime,
          peakInH,
          peakInHTime,
          peakInM,
          peakInMTime,
        }
      };
    }

    // =========================================================================
    // 7. RENDERIZADO DE UI PRINCIPAL Y GRÁFICOS CHART.JS
    // =========================================================================
    function syncSettingsInputsUI() {
      document.getElementById('inputClosingTime').value = state.closingTime;
      document.getElementById('selectBucketMinutes').value = String(state.bucketMinutes);
      document.getElementById('closingTimeBadge').textContent = `Cierre ${state.closingTime}`;
      document.getElementById('reportMetaSubtitle').textContent = `Sala ${state.roomCode} · Franjas de ${state.bucketMinutes} min`;
      document.getElementById('chartPeaksSubtitle').textContent = `Cantidad de ingresos cada ${state.bucketMinutes} minutos`;

      const b15 = document.getElementById('btnBucket15');
      const b30 = document.getElementById('btnBucket30');
      if (state.bucketMinutes === 15) {
        b15.className = 'no-select px-2 py-1 rounded-md bg-zinc-700 text-white';
        b30.className = 'no-select px-2 py-1 rounded-md text-zinc-400';
      } else {
        b30.className = 'no-select px-2 py-1 rounded-md bg-zinc-700 text-white';
        b15.className = 'no-select px-2 py-1 rounded-md text-zinc-400';
      }

      const lockedBanner = document.getElementById('lockedBanner');
      if (state.isLocked) {
        lockedBanner.classList.remove('hidden');
        document.getElementById('lockedBannerText').textContent = `Noche cerrada (${state.closingTime}) · Contador bloqueado`;
      } else {
        lockedBanner.classList.add('hidden');
      }
    }

    function recalculateAndRender() {
      const t = computeTotals();

      // Panel superior
      document.getElementById('statTotalIn').textContent = t.totalIn;
      document.getElementById('statTotalOut').textContent = t.totalOut;
      document.getElementById('statInsideTotal').textContent = t.insideTotal;
      document.getElementById('statInsideH').textContent = t.insideH;
      document.getElementById('statInsideM').textContent = t.insideM;
      document.getElementById('statPctH').textContent = `(${t.pctH}%)`;
      document.getElementById('statPctM').textContent = `(${t.pctM}%)`;

      const barH = document.getElementById('barH');
      const barM = document.getElementById('barM');
      if (t.insideTotal === 0) {
        barH.style.width = '50%';
        barM.style.width = '50%';
      } else {
        barH.style.width = `${t.pctH}%`;
        barM.style.width = `${t.pctM}%`;
      }

      // Contadores individuales en la botonera
      document.getElementById('countInH').textContent = t.inH;
      document.getElementById('countInM').textContent = t.inM;
      document.getElementById('countOutH').textContent = t.outH;
      document.getElementById('countOutM').textContent = t.outM;

      // Si estamos en la vista de reporte, actualizar gráficos y KPIs
      if (state.activeView === 'report') {
        renderReport();
      }
    }

    function getCommonChartOptions(yTitle, extraTooltipFn = null) {
      return {
        responsive: true,
        maintainAspectRatio: false,
        interaction: {
          mode: 'index',
          intersect: false,
        },
        plugins: {
          legend: {
            position: 'top',
            labels: {
              color: '#d4d4d8',
              usePointStyle: true,
              boxWidth: 8,
              font: { family: 'Inter', size: 11, weight: '600' },
            }
          },
          tooltip: {
            backgroundColor: 'rgba(17, 19, 24, 0.96)',
            titleColor: '#ffffff',
            bodyColor: '#e4e4e7',
            borderColor: '#3f3f46',
            borderWidth: 1,
            padding: 10,
            callbacks: extraTooltipFn ? { afterBody: extraTooltipFn } : {},
          }
        },
        scales: {
          x: {
            grid: { color: 'rgba(63, 63, 70, 0.25)' },
            ticks: {
              color: '#a1a1aa',
              font: { family: 'Inter', size: 10, weight: '600' },
              maxRotation: 45,
              autoSkip: true,
              maxTicksLimit: 10,
            }
          },
          y: {
            beginAtZero: true,
            title: {
              display: true,
              text: yTitle,
              color: '#71717a',
              font: { family: 'Inter', size: 10, weight: '700' },
            },
            grid: { color: 'rgba(63, 63, 70, 0.25)' },
            ticks: {
              color: '#a1a1aa',
              precision: 0,
              font: { family: 'Inter', size: 10, weight: '600' },
            }
          }
        }
      };
    }

    function renderReport() {
      const totals = computeTotals();
      const { buckets, kpis } = buildTimeSeries();

      // KPIs del reporte
      document.getElementById('kpiPeakInside').textContent = kpis.peakInside;
      document.getElementById('kpiPeakInsideTime').textContent = `Horario: ${kpis.peakInsideTime}`;
      document.getElementById('kpiTotalEntries').textContent = totals.totalIn;
      document.getElementById('kpiTotalSplit').innerHTML =
        `<span class="text-sky-400">${totals.inH} H (${totals.pctInH}%)</span> · <span class="text-pink-400">${totals.inM} M (${totals.pctInM}%)</span>`;
      document.getElementById('kpiPeakInM').textContent = `${kpis.peakInM} / franja`;
      document.getElementById('kpiPeakInMTime').textContent = `Horario: ${kpis.peakInMTime}`;
      document.getElementById('kpiPeakInH').textContent = `${kpis.peakInH} / franja`;
      document.getElementById('kpiPeakInHTime').textContent = `Horario: ${kpis.peakInHTime}`;

      const emptyEl = document.getElementById('reportEmptyState');
      const chartsEl = document.getElementById('chartsContainer');

      if (buckets.length === 0) {
        emptyEl.classList.remove('hidden');
        chartsEl.classList.add('hidden');
        return;
      }

      emptyEl.classList.add('hidden');
      chartsEl.classList.remove('hidden');

      const labels = buckets.map((b) => b.label);

      // Destruir instancias previas para evitar solapamientos
      if (charts.occupancy) charts.occupancy.destroy();
      if (charts.peaks) charts.peaks.destroy();
      if (charts.cumulative) charts.cumulative.destroy();

      // 1. GRÁFICO DE OCUPACIÓN Y PROPORCIÓN
      const occDatasets = [
        {
          label: 'Mujeres',
          data: buckets.map((b) => b.insideM),
          borderColor: '#ec4899',
          backgroundColor: 'rgba(236, 72, 153, 0.14)',
          borderWidth: 2.8,
          tension: 0.35,
          fill: true,
          pointRadius: 2.5,
          pointHoverRadius: 5,
        },
        {
          label: 'Hombres',
          data: buckets.map((b) => b.insideH),
          borderColor: '#38bdf8',
          backgroundColor: 'rgba(56, 189, 248, 0.14)',
          borderWidth: 2.8,
          tension: 0.35,
          fill: true,
          pointRadius: 2.5,
          pointHoverRadius: 5,
        },
      ];

      if (state.showTotalCurve) {
        occDatasets.push({
          label: 'Total Adentro',
          data: buckets.map((b) => b.insideTotal),
          borderColor: '#fbbf24',
          backgroundColor: 'transparent',
          borderWidth: 2,
          borderDash: [5, 4],
          tension: 0.35,
          fill: false,
          pointRadius: 0,
          pointHoverRadius: 4,
        });
      }

      charts.occupancy = new Chart(document.getElementById('chartOccupancy'), {
        type: 'line',
        data: { labels, datasets: occDatasets },
        options: getCommonChartOptions('Personas adentro', (items) => {
          const idx = items[0]?.dataIndex;
          if (idx === undefined || !buckets[idx]) return [];
          const b = buckets[idx];
          return [`Proporción: ${b.pctM}% Mujeres / ${b.pctH}% Hombres`];
        }),
      });

      // 2. GRÁFICO DE PICOS DE INGRESO POR FRANJA
      charts.peaks = new Chart(document.getElementById('chartPeaks'), {
        type: 'line',
        data: {
          labels,
          datasets: [
            {
              label: 'Ingresos Mujeres',
              data: buckets.map((b) => b.inM),
              borderColor: '#ec4899',
              backgroundColor: 'rgba(236, 72, 153, 0.18)',
              borderWidth: 2.8,
              tension: 0.35,
              fill: true,
              pointRadius: 3,
              pointHoverRadius: 5,
            },
            {
              label: 'Ingresos Hombres',
              data: buckets.map((b) => b.inH),
              borderColor: '#38bdf8',
              backgroundColor: 'rgba(56, 189, 248, 0.18)',
              borderWidth: 2.8,
              tension: 0.35,
              fill: true,
              pointRadius: 3,
              pointHoverRadius: 5,
            },
          ],
        },
        options: getCommonChartOptions(`Ingresos / ${state.bucketMinutes} min`),
      });

      // 3. GRÁFICO DE INGRESOS ACUMULADOS
      charts.cumulative = new Chart(document.getElementById('chartCumulative'), {
        type: 'line',
        data: {
          labels,
          datasets: [
            {
              label: 'Acumulado Mujeres',
              data: buckets.map((b) => b.cumInM),
              borderColor: '#ec4899',
              backgroundColor: 'rgba(236, 72, 153, 0.10)',
              borderWidth: 2.6,
              tension: 0.3,
              fill: true,
              pointRadius: 2,
            },
            {
              label: 'Acumulado Hombres',
              data: buckets.map((b) => b.cumInH),
              borderColor: '#38bdf8',
              backgroundColor: 'rgba(56, 189, 248, 0.10)',
              borderWidth: 2.6,
              tension: 0.3,
              fill: true,
              pointRadius: 2,
            },
          ],
        },
        options: getCommonChartOptions('Ingresos acumulados'),
      });
    }

    // =========================================================================
    // 8. EXPORTACIÓN PNG DE ALTA RESOLUCIÓN PARA WHATSAPP
    // =========================================================================
    function drawRoundRect(ctx, x, y, w, h, r) {
      ctx.beginPath();
      ctx.moveTo(x + r, y);
      ctx.arcTo(x + w, y, x + w, y + h, r);
      ctx.arcTo(x + w, y + h, x, y + h, r);
      ctx.arcTo(x, y + h, x, y, r);
      ctx.arcTo(x, y, x + w, y, r);
      ctx.closePath();
    }

    async function renderOffscreenChart(width, height, config) {
      const c = document.createElement('canvas');
      c.width = width;
      c.height = height;
      const ctx = c.getContext('2d');
      const instance = new Chart(ctx, {
        ...config,
        options: {
          ...config.options,
          responsive: false,
          animation: false,
          devicePixelRatio: 1,
        }
      });
      instance.update('none');
      const dataUrl = c.toDataURL('image/png');
      instance.destroy();

      return new Promise((resolve) => {
        const img = new Image();
        img.onload = () => resolve(img);
        img.src = dataUrl;
      });
    }

    async function generateReportPngBlob() {
      const totals = computeTotals();
      const { buckets, kpis } = buildTimeSeries();

      if (buckets.length === 0) {
        showToast('No hay datos todavía para generar la imagen', 'warn');
        return null;
      }

      const W = 1080;
      const H = 1760;
      const canvas = document.createElement('canvas');
      canvas.width = W;
      canvas.height = H;
      const ctx = canvas.getContext('2d');

      // Fondo AMOLED oscuro sólido
      ctx.fillStyle = '#09090b';
      ctx.fillRect(0, 0, W, H);

      // Franja de acento superior
      const grad = ctx.createLinearGradient(0, 0, W, 0);
      grad.addColorStop(0, '#38bdf8');
      grad.addColorStop(1, '#ec4899');
      ctx.fillStyle = grad;
      ctx.fillRect(0, 0, W, 8);

      // Encabezado
      const pad = 48;
      ctx.fillStyle = '#ffffff';
      ctx.font = '900 42px Inter, sans-serif';
      ctx.fillText('REPORTE DE PUERTA · NOCHE', pad, 76);

      const fechaStr = new Date().toLocaleDateString('es-AR', {
        weekday: 'long', day: '2-digit', month: '2-digit', year: 'numeric'
      });
      const rangoStr = `${buckets[0].label} a ${buckets[buckets.length - 1].label}`;
      ctx.fillStyle = '#a1a1aa';
      ctx.font = '600 24px Inter, sans-serif';
      ctx.fillText(`${fechaStr.toUpperCase()}  ·  SALA ${state.roomCode}  ·  ${rangoStr}`, pad, 114);

      // Tarjetas KPI (2x2)
      const cardW = (W - pad * 2 - 24) / 2;
      const cardH = 148;
      const kpiCards = [
        {
          title: 'TOTAL INGRESOS',
          value: String(totals.totalIn),
          sub: `${totals.inM} Mujeres (${totals.pctInM}%) · ${totals.inH} Hombres (${totals.pctInH}%)`,
          accent: '#ffffff',
        },
        {
          title: 'PICO MÁXIMO ADENTRO',
          value: `${kpis.peakInside} personas`,
          sub: `Registrado a las ${kpis.peakInsideTime} hs`,
          accent: '#fbbf24',
        },
        {
          title: 'PICO INGRESO MUJERES',
          value: `${kpis.peakInM} / ${state.bucketMinutes} min`,
          sub: `Franja pico: ${kpis.peakInMTime} hs`,
          accent: '#ec4899',
        },
        {
          title: 'PICO INGRESO HOMBRES',
          value: `${kpis.peakInH} / ${state.bucketMinutes} min`,
          sub: `Franja pico: ${kpis.peakInHTime} hs`,
          accent: '#38bdf8',
        },
      ];

      kpiCards.forEach((card, i) => {
        const col = i % 2;
        const row = Math.floor(i / 2);
        const x = pad + col * (cardW + 24);
        const y = 148 + row * (cardH + 20);

        drawRoundRect(ctx, x, y, cardW, cardH, 20);
        ctx.fillStyle = '#111318';
        ctx.fill();
        ctx.lineWidth = 2;
        ctx.strokeStyle = '#27272a';
        ctx.stroke();

        ctx.fillStyle = '#a1a1aa';
        ctx.font = '700 18px Inter, sans-serif';
        ctx.fillText(card.title, x + 24, y + 38);

        ctx.fillStyle = card.accent;
        ctx.font = '900 42px Inter, sans-serif';
        ctx.fillText(card.value, x + 24, y + 92);

        ctx.fillStyle = '#d4d4d8';
        ctx.font = '600 19px Inter, sans-serif';
        ctx.fillText(card.sub, x + 24, y + 126);
      });

      // Barra de proporción total de ingresos
      const barY = 486;
      const barW = W - pad * 2;
      drawRoundRect(ctx, pad, barY, barW, 24, 12);
      ctx.fillStyle = '#27272a';
      ctx.fill();

      const wH = Math.max(0, (totals.pctInH / 100) * barW);
      if (wH > 0) {
        ctx.save();
        drawRoundRect(ctx, pad, barY, barW, 24, 12);
        ctx.clip();
        ctx.fillStyle = '#38bdf8';
        ctx.fillRect(pad, barY, wH, 24);
        ctx.fillStyle = '#ec4899';
        ctx.fillRect(pad + wH, barY, barW - wH, 24);
        ctx.restore();
      }

      // Render offscreen de ambos gráficos principales con tipografía nítida
      const labels = buckets.map((b) => b.label);
      const chartBoxW = W - pad * 2;
      const chartCanvasW = chartBoxW - 48;
      const chartCanvasH = 440;

      const exportChartOpts = (yLabel) => ({
        plugins: {
          legend: {
            position: 'top',
            labels: {
              color: '#e4e4e7',
              usePointStyle: true,
              font: { family: 'Inter', size: 20, weight: '700' },
              padding: 20,
            }
          }
        },
        scales: {
          x: {
            grid: { color: 'rgba(63,63,70,0.35)' },
            ticks: { color: '#a1a1aa', font: { family: 'Inter', size: 18, weight: '600' }, maxTicksLimit: 12 }
          },
          y: {
            beginAtZero: true,
            title: { display: true, text: yLabel, color: '#a1a1aa', font: { family: 'Inter', size: 18, weight: '700' } },
            grid: { color: 'rgba(63,63,70,0.35)' },
            ticks: { color: '#a1a1aa', precision: 0, font: { family: 'Inter', size: 18, weight: '600' } }
          }
        }
      });

      const imgOcc = await renderOffscreenChart(chartCanvasW, chartCanvasH, {
        type: 'line',
        data: {
          labels,
          datasets: [
            {
              label: 'Mujeres adentro',
              data: buckets.map((b) => b.insideM),
              borderColor: '#ec4899',
              backgroundColor: 'rgba(236, 72, 153, 0.16)',
              borderWidth: 4,
              tension: 0.35,
              fill: true,
              pointRadius: 4,
            },
            {
              label: 'Hombres adentro',
              data: buckets.map((b) => b.insideH),
              borderColor: '#38bdf8',
              backgroundColor: 'rgba(56, 189, 248, 0.16)',
              borderWidth: 4,
              tension: 0.35,
              fill: true,
              pointRadius: 4,
            },
            {
              label: 'Total adentro',
              data: buckets.map((b) => b.insideTotal),
              borderColor: '#fbbf24',
              borderWidth: 3,
              borderDash: [8, 6],
              tension: 0.35,
              fill: false,
              pointRadius: 0,
            }
          ]
        },
        options: exportChartOpts('Personas adentro'),
      });

      const imgPeaks = await renderOffscreenChart(chartCanvasW, chartCanvasH, {
        type: 'line',
        data: {
          labels,
          datasets: [
            {
              label: 'Ingresos Mujeres',
              data: buckets.map((b) => b.inM),
              borderColor: '#ec4899',
              backgroundColor: 'rgba(236, 72, 153, 0.2)',
              borderWidth: 4,
              tension: 0.35,
              fill: true,
              pointRadius: 4,
            },
            {
              label: 'Ingresos Hombres',
              data: buckets.map((b) => b.inH),
              borderColor: '#38bdf8',
              backgroundColor: 'rgba(56, 189, 248, 0.2)',
              borderWidth: 4,
              tension: 0.35,
              fill: true,
              pointRadius: 4,
            }
          ]
        },
        options: exportChartOpts(`Ingresos / ${state.bucketMinutes} min`),
      });

      // Card Gráfico 1
      const g1Y = 540;
      const gCardH = 545;
      drawRoundRect(ctx, pad, g1Y, chartBoxW, gCardH, 24);
      ctx.fillStyle = '#111318';
      ctx.fill();
      ctx.strokeStyle = '#27272a';
      ctx.lineWidth = 2;
      ctx.stroke();

      ctx.fillStyle = '#ffffff';
      ctx.font = '800 24px Inter, sans-serif';
      ctx.fillText('1. OCUPACIÓN Y PROPORCIÓN A LO LARGO DE LA NOCHE', pad + 24, g1Y + 46);
      ctx.drawImage(imgOcc, pad + 24, g1Y + 72, chartCanvasW, chartCanvasH);

      // Card Gráfico 2
      const g2Y = g1Y + gCardH + 28;
      drawRoundRect(ctx, pad, g2Y, chartBoxW, gCardH, 24);
      ctx.fillStyle = '#111318';
      ctx.fill();
      ctx.strokeStyle = '#27272a';
      ctx.lineWidth = 2;
      ctx.stroke();

      ctx.fillStyle = '#ffffff';
      ctx.font = '800 24px Inter, sans-serif';
      ctx.fillText(`2. PICOS DE INGRESO POR FRANJA (${state.bucketMinutes} MIN)`, pad + 24, g2Y + 46);
      ctx.drawImage(imgPeaks, pad + 24, g2Y + 72, chartCanvasW, chartCanvasH);

      // Pie de imagen
      ctx.fillStyle = '#71717a';
      ctx.font = '600 20px Inter, sans-serif';
      ctx.fillText(
        `Total Salidas: ${totals.totalOut} (${totals.outH} H / ${totals.outM} M)  ·  Adentro al corte: ${totals.insideTotal}`,
        pad,
        H - 36
      );

      return new Promise((resolve) => {
        canvas.toBlob((blob) => resolve(blob), 'image/png', 1.0);
      });
    }

    function downloadBlob(blob, filename) {
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 2000);
    }

    async function handleExportPng(tryNativeShare = true) {
      showToast('Generando imagen de alta resolución...', 'info');
      try {
        const blob = await generateReportPngBlob();
        if (!blob) return;

        const filename = `reporte-boliche-${state.roomCode}-${formatDateYYYYMMDD()}.png`;
        const file = new File([blob], filename, { type: 'image/png' });
        const totals = computeTotals();

        if (tryNativeShare && navigator.canShare && navigator.canShare({ files: [file] })) {
          try {
            await navigator.share({
              files: [file],
              title: `Reporte de Puerta · Sala ${state.roomCode}`,
              text: `Resumen de la noche: ${totals.totalIn} ingresos (${totals.inM} M / ${totals.inH} H).`,
            });
            showToast('Imagen lista para compartir', 'success');
            return;
          } catch (shareErr) {
            if (shareErr.name === 'AbortError') return;
          }
        }

        // Respaldo: descarga directa PNG
        downloadBlob(blob, filename);
        showToast('Imagen PNG descargada. ¡Listo para mandar por WhatsApp!', 'success');
      } catch (err) {
        console.error('Error generando PNG:', err);
        showToast('No se pudo generar la imagen PNG', 'error');
      }
    }

    // =========================================================================
    // 9. EXPORTACIÓN A EXCEL (.xlsx) CON 3 HOJAS (SHEETJS)
    // =========================================================================
    function handleExportExcel() {
      const active = getActiveEvents();
      if (active.length === 0) {
        showToast('No hay registros todavía para exportar a Excel', 'warn');
        return;
      }

      if (typeof XLSX === 'undefined') {
        showToast('Librería de Excel no disponible', 'error');
        return;
      }

      try {
        const totals = computeTotals();
        const { buckets, kpis } = buildTimeSeries();
        const wb = XLSX.utils.book_new();

        // HOJA 1: Resumen
        const summaryRows = [
          ['REPORTE DE CONTROL DE PUERTA', ''],
          ['Fecha de exportación', new Date().toLocaleString('es-AR')],
          ['Código de Sala', state.roomCode],
          ['Hora de cierre configurada', state.closingTime],
          ['Intervalo de gráficos', `${state.bucketMinutes} minutos`],
          ['', ''],
          ['MÉTRICA', 'VALOR'],
          ['Total Ingresos', totals.totalIn],
          ['Ingresos Hombres', `${totals.inH} (${totals.pctInH}%)`],
          ['Ingresos Mujeres', `${totals.inM} (${totals.pctInM}%)`],
          ['Total Salidas', totals.totalOut],
          ['Salidas Hombres', totals.outH],
          ['Salidas Mujeres', totals.outM],
          ['Adentro Actual (Total)', totals.insideTotal],
          ['Adentro Actual Hombres', `${totals.insideH} (${totals.pctH}%)`],
          ['Adentro Actual Mujeres', `${totals.insideM} (${totals.pctM}%)`],
          ['', ''],
          ['PICOS DE LA NOCHE', ''],
          ['Pico Máximo Adentro', kpis.peakInside],
          ['Horario Pico Máximo Adentro', kpis.peakInsideTime],
          ['Pico Ingreso Mujeres (por franja)', kpis.peakInM],
          ['Horario Pico Ingreso Mujeres', kpis.peakInMTime],
          ['Pico Ingreso Hombres (por franja)', kpis.peakInH],
          ['Horario Pico Ingreso Hombres', kpis.peakInHTime],
        ];
        const wsSummary = XLSX.utils.aoa_to_sheet(summaryRows);
        wsSummary['!cols'] = [{ wch: 34 }, { wch: 28 }];
        XLSX.utils.book_append_sheet(wb, wsSummary, 'Resumen');

        // HOJA 2: Evolución por Franja
        const evolutionRows = [
          [
            'Horario',
            'Ingresos Hombres',
            'Ingresos Mujeres',
            'Total Ingresos Franja',
            'Salidas Hombres',
            'Salidas Mujeres',
            'Adentro Hombres',
            'Adentro Mujeres',
            'Total Adentro',
            '% Hombres',
            '% Mujeres',
            'Acumulado Hombres',
            'Acumulado Mujeres',
          ],
          ...buckets.map((b) => [
            b.label,
            b.inH,
            b.inM,
            b.inH + b.inM,
            b.outH,
            b.outM,
            b.insideH,
            b.insideM,
            b.insideTotal,
            `${b.pctH}%`,
            `${b.pctM}%`,
            b.cumInH,
            b.cumInM,
          ]),
        ];
        const wsEvolution = XLSX.utils.aoa_to_sheet(evolutionRows);
        wsEvolution['!cols'] = Array(13).fill({ wch: 18 });
        XLSX.utils.book_append_sheet(wb, wsEvolution, 'Evolución por Franja');

        // HOJA 3: Registro Detallado
        const detailRows = [
          ['Fecha', 'Hora Exacta', 'Tipo', 'Género', 'Cantidad (Delta)', 'Origen / Operador'],
          ...active.map((ev) => {
            const d = new Date(ev.timestamp);
            return [
              d.toLocaleDateString('es-AR'),
              d.toLocaleTimeString('es-AR', { hour12: false }),
              ev.tipo === 'in' ? 'Ingreso' : 'Salida',
              ev.genero === 'H' ? 'Hombre' : 'Mujer',
              ev.delta,
              ev.simulated ? 'Simulación' : (ev.uid || 'Operador'),
            ];
          }),
        ];
        const wsDetail = XLSX.utils.aoa_to_sheet(detailRows);
        wsDetail['!cols'] = [{ wch: 14 }, { wch: 14 }, { wch: 12 }, { wch: 12 }, { wch: 16 }, { wch: 28 }];
        XLSX.utils.book_append_sheet(wb, wsDetail, 'Registro Detallado');

        const filename = `reporte-boliche-${state.roomCode}-${formatDateYYYYMMDD()}.xlsx`;
        XLSX.writeFile(wb, filename);
        showToast('Excel (.xlsx) exportado con 3 hojas', 'success');
      } catch (err) {
        console.error('Error exportando Excel:', err);
        showToast('Error al generar el archivo Excel', 'error');
      }
    }

    // =========================================================================
    // 10. SIMULACIÓN DE NOCHE DE PRUEBA Y REINICIO
    // =========================================================================
    async function loadSimulationNight() {
      closeModal('settingsModal');
      showToast('Generando noche de prueba (23:30 a 05:00)...', 'info');

      // Base temporal: armamos una noche terminando hoy a las 05:00 (o 5.5 hs atrás hasta ahora)
      const now = new Date();
      const baseStart = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 30, 0, 0);
      // Si todavía no pasaron las 23:30, usamos ayer 23:30 -> hoy 05:00
      if (now.getHours() < 23) {
        baseStart.setDate(baseStart.getDate() - 1);
      }

      // Curva realista de boliche en intervalos de 15 min (23:30 a 05:00 = 23 franjas)
      const profile = [
        // [minOffset, inM, inH, outM, outH]
        [0,   10,  6,  0,  0],  // 23:30
        [15,  16, 10,  0,  0],  // 23:45
        [30,  22, 14,  0,  0],  // 00:00
        [45,  28, 19,  1,  0],  // 00:15
        [60,  36, 25,  1,  1],  // 00:30
        [75,  44, 31,  2,  1],  // 00:45
        [90,  52, 38,  2,  2],  // 01:00 -> Pico temprano mujeres
        [105, 48, 45,  3,  2],  // 01:15
        [120, 42, 54,  4,  3],  // 01:30 -> Pico fuertes hombres
        [135, 34, 49,  5,  4],  // 01:45
        [150, 26, 40,  7,  5],  // 02:00
        [165, 19, 30,  9,  7],  // 02:15
        [180, 14, 22, 12,  9],  // 02:30
        [195,  9, 15, 16, 12],  // 02:45
        [210,  6, 10, 22, 16],  // 03:00
        [225,  4,  6, 28, 22],  // 03:15
        [240,  2,  4, 34, 28],  // 03:30
        [255,  1,  2, 40, 35],  // 03:45
        [270,  0,  1, 46, 42],  // 04:00
        [285,  0,  0, 52, 49],  // 04:15
        [300,  0,  0, 48, 52],  // 04:30
        [315,  0,  0, 38, 46],  // 04:45
        [330,  0,  0, 26, 36],  // 05:00
      ];

      const simEvents = [];
      const startMs = baseStart.getTime();
      const resetEpoch = startMs - 60000;

      for (const [offsetMin, inM, inH, outM, outH] of profile) {
        const baseTs = startMs + offsetMin * 60 * 1000;
        if (inM > 0) simEvents.push({ tipo: 'in', genero: 'M', delta: inM, timestamp: baseTs + 10000, uid: 'sim', simulated: true });
        if (inH > 0) simEvents.push({ tipo: 'in', genero: 'H', delta: inH, timestamp: baseTs + 20000, uid: 'sim', simulated: true });
        if (outM > 0) simEvents.push({ tipo: 'out', genero: 'M', delta: outM, timestamp: baseTs + 30000, uid: 'sim', simulated: true });
        if (outH > 0) simEvents.push({ tipo: 'out', genero: 'H', delta: outH, timestamp: baseTs + 40000, uid: 'sim', simulated: true });
      }

      if (!firebaseReady || !currentUser) {
        state.resetEpoch = resetEpoch;
        state.isLocked = false;
        state.events = simEvents.map((ev, i) => ({ id: `sim_${i}`, ...ev }));
        recalculateAndRender();
        openView('report');
        showToast('Noche de prueba cargada', 'success');
        return;
      }

      try {
        // Limpiar documentos existentes en la colección de la sala
        const eventsColRef = collection(db, 'artifacts', appId, 'public', 'data', `club_events_${state.roomCode}`);
        const existing = await getDocs(eventsColRef);
        await Promise.all(existing.docs.map((d) => deleteDoc(d.ref)));

        await updateRoomMeta({ resetEpoch, isLocked: false });

        // Cargar eventos simulados en paralelo en bloques pequeños
        const chunkSize = 15;
        for (let i = 0; i < simEvents.length; i += chunkSize) {
          const chunk = simEvents.slice(i, i + chunkSize);
          await Promise.all(chunk.map((ev) => addDoc(eventsColRef, ev)));
        }

        openView('report');
        showToast('Noche de prueba sincronizada y lista', 'success');
      } catch (err) {
        console.error('Error cargando simulación:', err);
        showToast('Error al cargar simulación', 'error');
      }
    }

    async function resetNight() {
      const nowEpoch = Date.now();
      document.getElementById('resetConfirmBox').classList.add('hidden');
      document.getElementById('btnResetStep1').classList.remove('hidden');
      closeModal('settingsModal');

      if (!firebaseReady || !currentUser) {
        state.events = [];
        state.resetEpoch = nowEpoch;
        state.isLocked = false;
        recalculateAndRender();
        showToast('Contadores reiniciados a 0', 'success');
        return;
      }

      try {
        // Actualización inmediata del resetEpoch para limpiar al instante todos los celulares
        await updateRoomMeta({
          resetEpoch: nowEpoch,
          isLocked: false,
          lastAutoClosedKey: '',
        });

        // Limpieza en segundo plano de documentos anteriores
        const eventsColRef = collection(db, 'artifacts', appId, 'public', 'data', `club_events_${state.roomCode}`);
        const existing = await getDocs(eventsColRef);
        const toDelete = existing.docs.filter((d) => (Number(d.data().timestamp) || 0) < nowEpoch);
        Promise.all(toDelete.map((d) => deleteDoc(d.ref))).catch(() => {});

        document.getElementById('lastActionText').textContent = 'Noche reiniciada';
        showToast('Noche reiniciada en todos los dispositivos', 'success');
      } catch (err) {
        console.error('Error reiniciando noche:', err);
        showToast('No se pudo reiniciar la noche', 'error');
      }
    }

    // =========================================================================
    // 11. RELOJ EN VIVO Y CIERRE AUTOMÁTICO A LAS 05:00 AM (CONFIGURABLE)
    // =========================================================================
    function startClockWatcher() {
      const tick = () => {
        const now = new Date();
        const hhmm = formatTimeHHMM(now.getTime());
        document.getElementById('liveClock').textContent = hhmm;
        document.getElementById('mobileClockRow').textContent = `Hora: ${hhmm} · Cierre auto: ${state.closingTime}`;

        // Verificar si llegó el horario exacto de cierre automático
        const todayKey = `${formatDateYYYYMMDD(now)} ${state.closingTime}`;
        if (
          hhmm === state.closingTime &&
          state.lastAutoClosedKey !== todayKey &&
          !state.isLocked
        ) {
          state.lastAutoClosedKey = todayKey;
          updateRoomMeta({
            isLocked: true,
            lastAutoClosedKey: todayKey,
          });
          vibrate([60, 40, 60]);
          openView('report');
          showToast(`Cierre automático (${state.closingTime}) activado · Reporte generado`, 'warn');
        }
      };

      tick();
      setInterval(tick, 1000);
    }

    // =========================================================================
    // 12. NAVEGACIÓN DE VISTAS Y MODALES
    // =========================================================================
    function openView(viewName) {
      state.activeView = viewName;
      const counterView = document.getElementById('counterView');
      const reportView = document.getElementById('reportView');

      if (viewName === 'report') {
        counterView.classList.add('hidden');
        reportView.classList.remove('hidden');
        reportView.classList.add('flex');
        renderReport();
        window.scrollTo({ top: 0, behavior: 'smooth' });
      } else {
        reportView.classList.add('hidden');
        reportView.classList.remove('flex');
        counterView.classList.remove('hidden');
      }
    }

    function openModal(id) {
      document.getElementById(id).classList.remove('hidden');
    }

    function closeModal(id) {
      document.getElementById(id).classList.add('hidden');
    }

    // =========================================================================
    // 13. LISTENERS DE EVENTOS
    // =========================================================================
    function bindEvents() {
      // Botones principales de conteo (+1)
      const btnInH = document.getElementById('btnInH');
      const btnInM = document.getElementById('btnInM');
      const btnOutH = document.getElementById('btnOutH');
      const btnOutM = document.getElementById('btnOutM');

      btnInH.addEventListener('click', () => recordEvent('in', 'H', 1, btnInH));
      btnInM.addEventListener('click', () => recordEvent('in', 'M', 1, btnInM));
      btnOutH.addEventListener('click', () => recordEvent('out', 'H', 1, btnOutH));
      btnOutM.addEventListener('click', () => recordEvent('out', 'M', 1, btnOutM));

      // Botones chicos de corrección (-1)
      document.getElementById('btnUndoInH').addEventListener('click', (e) => {
        e.stopPropagation();
        recordEvent('in', 'H', -1);
      });
      document.getElementById('btnUndoInM').addEventListener('click', (e) => {
        e.stopPropagation();
        recordEvent('in', 'M', -1);
      });
      document.getElementById('btnUndoOutH').addEventListener('click', (e) => {
        e.stopPropagation();
        recordEvent('out', 'H', -1);
      });
      document.getElementById('btnUndoOutM').addEventListener('click', (e) => {
        e.stopPropagation();
        recordEvent('out', 'M', -1);
      });

      // Desbloquear contador si se disparó el cierre automático
      document.getElementById('btnUnlockCounter').addEventListener('click', () => {
        updateRoomMeta({ isLocked: false });
        showToast('Contador desbloqueado', 'info');
      });

      // Navegación entre Contador y Reporte
      document.getElementById('btnOpenReport').addEventListener('click', () => {
        openView(state.activeView === 'report' ? 'counter' : 'report');
      });
      document.getElementById('btnBackToCounter').addEventListener('click', () => openView('counter'));

      // Selector de curva total e intervalos en el reporte
      document.getElementById('toggleShowTotalCurve').addEventListener('change', (e) => {
        state.showTotalCurve = e.target.checked;
        renderReport();
      });

      document.getElementById('btnBucket15').addEventListener('click', () => {
        updateRoomMeta({ bucketMinutes: 15 });
      });
      document.getElementById('btnBucket30').addEventListener('click', () => {
        updateRoomMeta({ bucketMinutes: 30 });
      });

      // Exportaciones
      document.getElementById('btnShareWhatsApp').addEventListener('click', () => handleExportPng(true));
      document.getElementById('btnDownloadPng').addEventListener('click', () => handleExportPng(false));
      document.getElementById('btnExportExcel').addEventListener('click', handleExportExcel);

      // Modal de Sincronización
      document.getElementById('btnOpenSync').addEventListener('click', () => openModal('syncModal'));
      document.getElementById('btnCloseSync').addEventListener('click', () => closeModal('syncModal'));
      document.getElementById('syncModal').addEventListener('click', (e) => {
        if (e.target.id === 'syncModal') closeModal('syncModal');
      });

      document.getElementById('btnCopyCode').addEventListener('click', () => {
        copyTextFallback(state.roomCode);
        showToast(`Código ${state.roomCode} copiado`, 'success');
      });

      document.getElementById('btnCopyLink').addEventListener('click', () => {
        copyTextFallback(getInviteUrl());
        showToast('Enlace de invitación copiado', 'success');
      });

      document.getElementById('btnJoinRoom').addEventListener('click', async () => {
        const input = document.getElementById('inputJoinCode');
        const code = input.value.trim().toUpperCase();
        if (!code || code.length < 4) {
          showToast('Ingresá un código de sala válido', 'warn');
          return;
        }
        input.value = '';
        closeModal('syncModal');
        await connectToRoom(code);
        showToast(`Conectado a la sala ${code}`, 'success');
      });

      document.getElementById('btnCreateNewRoom').addEventListener('click', async () => {
        const newCode = generateRoomCode();
        closeModal('syncModal');
        await connectToRoom(newCode);
        showToast(`Nueva sala creada: ${newCode}`, 'success');
      });

      // Modal de Configuración
      document.getElementById('btnOpenSettings').addEventListener('click', () => {
        document.getElementById('resetConfirmBox').classList.add('hidden');
        document.getElementById('btnResetStep1').classList.remove('hidden');
        openModal('settingsModal');
      });
      document.getElementById('btnCloseSettings').addEventListener('click', () => closeModal('settingsModal'));
      document.getElementById('settingsModal').addEventListener('click', (e) => {
        if (e.target.id === 'settingsModal') closeModal('settingsModal');
      });

      document.getElementById('inputClosingTime').addEventListener('change', (e) => {
        const val = e.target.value || '05:00';
        updateRoomMeta({ closingTime: val, lastAutoClosedKey: '' });
        showToast(`Cierre automático configurado a las ${val}`, 'success');
      });

      document.getElementById('selectBucketMinutes').addEventListener('change', (e) => {
        const val = Number(e.target.value) === 30 ? 30 : 15;
        updateRoomMeta({ bucketMinutes: val });
      });

      // Simulación
      document.getElementById('btnLoadSimulation').addEventListener('click', loadSimulationNight);
      document.getElementById('btnEmptyLoadSim').addEventListener('click', loadSimulationNight);

      // Reinicio con doble confirmación
      document.getElementById('btnResetStep1').addEventListener('click', () => {
        document.getElementById('btnResetStep1').classList.add('hidden');
        document.getElementById('resetConfirmBox').classList.remove('hidden');
      });
      document.getElementById('btnResetCancel').addEventListener('click', () => {
        document.getElementById('resetConfirmBox').classList.add('hidden');
        document.getElementById('btnResetStep1').classList.remove('hidden');
      });
      document.getElementById('btnResetConfirm').addEventListener('click', resetNight);
    }

    // =========================================================================
    // 14. ARRANQUE DE LA APP
    // =========================================================================
    window.addEventListener('DOMContentLoaded', () => {
      bindEvents();
      startClockWatcher();
      initFirebase();
    });
  </script>
</body>
</html>
```
