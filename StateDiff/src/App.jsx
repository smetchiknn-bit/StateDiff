import { useState, useCallback } from 'react'
import * as XLSX from 'xlsx'

function App() {
  const [step, setStep] = useState(1)
  const [fileBylo, setFileBylo] = useState(null)
  const [fileStalo, setFileStalo] = useState(null)
  const [loading, setLoading] = useState(false)
  const [result, setResult] = useState(null)
  const [summary, setSummary] = useState(null)
  const [dragOver, setDragOver] = useState(false)
  const [fileNameBylo, setFileNameBylo] = useState('')
  const [fileNameStalo, setFileNameStalo] = useState('')

  // Обработка файлов Excel
  const processExcelFile = async (file, type) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader()
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result)
          const workbook = XLSX.read(data, { type: 'array' })
          
          const smrSheet = workbook.Sheets['СМР уник']
          const tmcSheet = workbook.Sheets['ТМЦ уник']
          
          if (!smrSheet || !tmcSheet) {
            reject(new Error(`В файле "${type}" нет листов "СМР уник" или "ТМЦ уник"`))
            return
          }
          
          const smrData = XLSX.utils.sheet_to_json(smrSheet, { header: 1 })
          const tmcData = XLSX.utils.sheet_to_json(tmcSheet, { header: 1 })
          
          resolve({
            smr: smrData.slice(1), // Пропускаем заголовок
            tmc: tmcData.slice(1)
          })
        } catch (err) {
          reject(err)
        }
      }
      reader.onerror = reject
      reader.readAsArrayBuffer(file)
    })
  }

  // Форматирование числа в русском формате
  const formatMoney = (num) => {
    if (num === null || num === undefined || isNaN(num)) return '0,00'
    return num.toLocaleString('ru-RU', { 
      minimumFractionDigits: 2, 
      maximumFractionDigits: 2 
    }) + ' руб.'
  }

  // Основная функция сравнения
  const compareFiles = async () => {
    if (!fileBylo || !fileStalo) return
    
    setLoading(true)
    try {
      const dataBylo = await processExcelFile(fileBylo, 'БЫЛО')
      const dataStalo = await processExcelFile(fileStalo, 'СТАЛО')
      
      // Создаём карту данных "БЫЛО" по кодам
      const byloMap = new Map()
      dataBylo.smr.forEach(row => {
        if (row[0]) {
          byloMap.set(`SMR_${row[0]}`, {
            code: row[0],
            name: row[1] || '',
            unit: row[2] || '',
            volume: row[3] || 0,
            price: row[4] || 0,
            type: 'SMR'
          })
        }
      })
      dataBylo.tmc.forEach(row => {
        if (row[0]) {
          byloMap.set(`TMC_${row[0]}`, {
            code: row[0],
            name: row[1] || '',
            unit: row[2] || '',
            volume: row[3] || 0,
            price: row[4] || 0,
            type: 'TMC'
          })
        }
      })
      
      // Обрабатываем данные "СТАЛО" и обновляем/добавляем позиции
      const staloKeys = new Set()
      
      dataStalo.smr.forEach(row => {
        if (row[0]) {
          const key = `SMR_${row[0]}`
          staloKeys.add(key)
          if (byloMap.has(key)) {
            const item = byloMap.get(key)
            item.staloVolume = row[3] || 0
            item.staloPrice = row[4] || 0
          } else {
            byloMap.set(key, {
              code: row[0],
              name: row[1] || '',
              unit: row[2] || '',
              volume: 0,
              price: 0,
              staloVolume: row[3] || 0,
              staloPrice: row[4] || 0,
              type: 'SMR',
              isNew: false
            })
          }
        }
      })
      
      dataStalo.tmc.forEach(row => {
        if (row[0]) {
          const key = `TMC_${row[0]}`
          staloKeys.add(key)
          if (byloMap.has(key)) {
            const item = byloMap.get(key)
            item.staloVolume = row[3] || 0
            item.staloPrice = row[4] || 0
          } else {
            byloMap.set(key, {
              code: row[0],
              name: row[1] || '',
              unit: row[2] || '',
              volume: 0,
              price: 0,
              staloVolume: row[3] || 0,
              staloPrice: row[4] || 0,
              type: 'TMC',
              isNew: false
            })
          }
        }
      })
      
      // Преобразуем в массив для отображения
      const resultData = Array.from(byloMap.values())
      
      // Считаем суммы
      let smrByloTotal = 0, smrStaloTotal = 0
      let tmcByloTotal = 0, tmcStaloTotal = 0
      
      resultData.forEach(item => {
        const byloSum = (item.volume || 0) * (item.price || 0)
        const staloSum = (item.staloVolume || 0) * (item.staloPrice || 0)
        
        if (item.type === 'SMR') {
          smrByloTotal += byloSum
          smrStaloTotal += staloSum
        } else {
          tmcByloTotal += byloSum
          tmcStaloTotal += staloSum
        }
        
        item.byloSum = Math.round(byloSum * 100) / 100
        item.staloSum = Math.round(staloSum * 100) / 100
        item.diff = Math.round((staloSum - byloSum) * 100) / 100
      })
      
      setResult(resultData)
      setSummary({
        smrBylo: Math.round(smrByloTotal * 100) / 100,
        smrStalo: Math.round(smrStaloTotal * 100) / 100,
        tmcBylo: Math.round(tmcByloTotal * 100) / 100,
        tmcStalo: Math.round(tmcStaloTotal * 100) / 100,
        totalBylo: Math.round((smrByloTotal + tmcByloTotal) * 100) / 100,
        totalStalo: Math.round((smrStaloTotal + tmcStaloTotal) * 100) / 100
      })
      
      setStep(4)
    } catch (err) {
      alert('Ошибка: ' + err.message)
    } finally {
      setLoading(false)
    }
  }

  // Обработчики drag & drop
  const handleDrop = useCallback((e, type) => {
    e.preventDefault()
    setDragOver(false)
    
    const files = e.dataTransfer.files
    if (files.length > 0) {
      const file = files[0]
      if (type === 'bylo') {
        setFileBylo(file)
        setFileNameBylo(file.name)
      } else {
        setFileStalo(file)
        setFileNameStalo(file.name)
      }
    }
  }, [])

  const handleFileSelect = (e, type) => {
    const files = e.target.files
    if (files.length > 0) {
      const file = files[0]
      if (type === 'bylo') {
        setFileBylo(file)
        setFileNameBylo(file.name)
      } else {
        setFileStalo(file)
        setFileNameStalo(file.name)
      }
    }
  }

  const handleReset = () => {
    setStep(1)
    setFileBylo(null)
    setFileStalo(null)
    setResult(null)
    setSummary(null)
    setFileNameBylo('')
    setFileNameStalo('')
  }

  // Рендер шагов прогресс-рейла
  const renderStep = (num, title, icon) => {
    let className = 'step-locked'
    if (step > num) className = 'step-done'
    else if (step === num) className = 'step-current'
    
    return (
      <div 
        className={`flex items-center px-4 py-2 rounded-lg cursor-pointer transition-all ${className}`}
        onClick={() => step > num && setStep(num)}
      >
        <div className="w-6 h-6 rounded-full flex items-center justify-center mr-2 text-sm font-bold">
          {step > num ? '✓' : num}
        </div>
        <span className="text-sm font-medium hidden lg:block">{title}</span>
      </div>
    )
  }

  return (
    <div className="min-h-screen relative">
      {/* Фоновые элементы */}
      <div className="bg-gradient-overlay"></div>
      
      {/* Плавающие символы */}
      <div className="float-symbol" style={{ top: '10%', right: '15%' }}>Σ</div>
      <div className="float-symbol" style={{ bottom: '20%', left: '10%', animationDelay: '3s' }}>₽</div>
      <div className="float-symbol" style={{ top: '40%', left: '5%', animationDelay: '6s' }}>=</div>
      
      {/* Шапка */}
      <header className="fixed top-0 left-0 right-0 bg-[#123f28] text-white z-50 shadow-lg">
        <div className="container mx-auto px-4 py-4">
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-4">
              {/* Логотип */}
              <div className="w-10 h-10 bg-[#1e7145] flex items-center justify-center shadow-hard-sm rounded">
                <img src="/logo.png" alt="Logo" className="w-8 h-8 object-contain" />
              </div>
              <div>
                <h1 className="font-unbounded text-lg font-semibold">StateDiff</h1>
                <p className="text-xs text-white/65">Сравнение сметных данных · БЫЛО / СТАЛО</p>
              </div>
            </div>
            
            {(fileBylo || fileStalo) && (
              <div className="flex items-center gap-4">
                <div className="hidden md:flex items-center gap-2 text-xs font-mono bg-black/20 px-3 py-1.5 rounded">
                  <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                  </svg>
                  {fileNameBylo && fileNameStalo ? `${fileNameBylo} / ${fileNameStalo}` : 'Файлы загружены'}
                </div>
                <button onClick={handleReset} className="btn-ghost text-white hover:border-red-500 hover:text-red-400">
                  Сброс
                </button>
              </div>
            )}
          </div>
          
          {/* Строка формул */}
          <div className="mt-3 pt-3 border-t border-white/10 flex items-center gap-3 text-xs">
            <span className="bg-[#1e7145] px-2 py-0.5 rounded font-mono">A1:J1000</span>
            <span className="text-white/60">ƒx</span>
            <span className="font-mono text-white/80">
              {step === 1 ? '=ожидание_файла_БЫЛО(...)' : 
               step === 2 ? '=ожидание_файла_СТАЛО(...)' :
               step === 3 ? '=расчёт_разницы(БЫЛО; СТАЛО)' :
               '=результат_готов'}
            </span>
          </div>
        </div>
      </header>
      
      {/* Прогресс-рейл */}
      <div className="fixed top-[120px] left-0 right-0 z-40 bg-[#eef2ec]/95 backdrop-blur-sm border-b border-[#cfd9d0]">
        <div className="container mx-auto px-4 py-3">
          <div className="flex items-center gap-2">
            {renderStep(1, 'Файл БЫЛО', '📁')}
            
            <div className={`flex-1 h-0.5 ${step >= 2 ? 'step-connector-done' : 'step-connector-future'}`}></div>
            
            {renderStep(2, 'Файл СТАЛО', '📁')}
            
            <div className={`flex-1 h-0.5 ${step >= 3 ? 'step-connector-done' : 'step-connector-future'}`}></div>
            
            {renderStep(3, 'Обработка', '⚙️')}
            
            <div className={`flex-1 h-0.5 ${step >= 4 ? 'step-connector-done' : 'step-connector-future'}`}></div>
            
            {renderStep(4, 'Результат', '📊')}
            
            {/* Индикатор безопасности */}
            <div className="hidden lg:flex items-center gap-2 ml-4 text-xs text-[#42544b]">
              <svg className="w-4 h-4 text-[#1e7145]" fill="currentColor" viewBox="0 0 20 20">
                <path fillRule="evenodd" d="M5 9V7a5 5 0 0110 0v2a2 2 0 012 2v5a2 2 0 01-2 2H5a2 2 0 01-2-2v-5a2 2 0 012-2zm8-2v2H7V7a3 3 0 016 0z" clipRule="evenodd" />
              </svg>
              <span>файл обрабатывается локально</span>
            </div>
          </div>
        </div>
      </div>
      
      {/* Основной контент */}
      <main className="container mx-auto px-4 pt-[180px] pb-8 relative z-10">
        <div className="grid lg:grid-cols-[1fr_400px] gap-6">
          {/* Левая колонка */}
          <div className="card p-6 rise-in">
            {/* Шаг 1: Загрузка файла БЫЛО */}
            {step === 1 && (
              <div>
                <div className="flex items-center gap-2 mb-4">
                  <svg className="w-5 h-5 text-[#1e7145]" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                  </svg>
                  <h2 className="font-unbounded text-lg font-semibold">Файл БЫЛО</h2>
                </div>
                
                {!fileBylo ? (
                  <div 
                    className={`dropzone rounded-xl p-8 text-center cursor-pointer transition-all ${dragOver ? 'drag-over' : ''}`}
                    onDragOver={(e) => { e.preventDefault(); setDragOver(true) }}
                    onDragLeave={() => setDragOver(false)}
                    onDrop={(e) => handleDrop(e, 'bylo')}
                    onClick={() => document.getElementById('file-bylo').click()}
                  >
                    <div className="w-14 h-14 mx-auto mb-4 bg-white border-2 border-[#1e7145] rounded flex items-center justify-center">
                      <svg className="w-6 h-6 text-[#1e7145]" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m0-3v12" />
                      </svg>
                    </div>
                    <h3 className="font-unbounded text-base font-semibold mb-2">Перетащите файл БЫЛО сюда</h3>
                    <p className="text-sm text-gray-500 mb-4">или нажмите, чтобы выбрать файл · .xlsx / .xlsm</p>
                    <p className="text-xs text-gray-400 max-w-md mx-auto">
                      Файл обрабатывается локально в браузере и никуда не отправляется.
                    </p>
                    <input 
                      id="file-bylo" 
                      type="file" 
                      accept=".xlsx,.xlsm,.xls" 
                      className="hidden"
                      onChange={(e) => handleFileSelect(e, 'bylo')}
                    />
                  </div>
                ) : (
                  <div className="bg-[#eef2ec]/70 border border-[#cfd9d0] rounded-lg p-4 flex items-center justify-between">
                    <div className="flex items-center gap-3">
                      <div className="w-8 h-8 bg-[#1e7145] rounded flex items-center justify-center">
                        <svg className="w-5 h-5 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                        </svg>
                      </div>
                      <div>
                        <p className="font-mono text-sm font-semibold">{fileNameBylo}</p>
                        <p className="text-xs text-gray-500">Готов к обработке</p>
                      </div>
                    </div>
                    <button 
                      onClick={() => { setFileBylo(null); setFileNameBylo('') }}
                      className="btn-ghost text-sm"
                    >
                      Заменить
                    </button>
                  </div>
                )}
                
                {fileBylo && (
                  <button 
                    onClick={() => setStep(2)}
                    className="btn-primary mt-6 w-full"
                  >
                    Далее: Файл СТАЛО →
                  </button>
                )}
              </div>
            )}
            
            {/* Шаг 2: Загрузка файла СТАЛО */}
            {step === 2 && (
              <div>
                <div className="flex items-center gap-2 mb-4">
                  <svg className="w-5 h-5 text-[#1e7145]" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                  </svg>
                  <h2 className="font-unbounded text-lg font-semibold">Файл СТАЛО</h2>
                </div>
                
                {!fileStalo ? (
                  <div 
                    className={`dropzone rounded-xl p-8 text-center cursor-pointer transition-all ${dragOver ? 'drag-over' : ''}`}
                    onDragOver={(e) => { e.preventDefault(); setDragOver(true) }}
                    onDragLeave={() => setDragOver(false)}
                    onDrop={(e) => handleDrop(e, 'stalo')}
                    onClick={() => document.getElementById('file-stalo').click()}
                  >
                    <div className="w-14 h-14 mx-auto mb-4 bg-white border-2 border-[#1e7145] rounded flex items-center justify-center">
                      <svg className="w-6 h-6 text-[#1e7145]" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m0-3v12" />
                      </svg>
                    </div>
                    <h3 className="font-unbounded text-base font-semibold mb-2">Перетащите файл СТАЛО сюда</h3>
                    <p className="text-sm text-gray-500 mb-4">или нажмите, чтобы выбрать файл · .xlsx / .xlsm</p>
                    <p className="text-xs text-gray-400 max-w-md mx-auto">
                      Файл обрабатывается локально в браузере и никуда не отправляется.
                    </p>
                    <input 
                      id="file-stalo" 
                      type="file" 
                      accept=".xlsx,.xlsm,.xls" 
                      className="hidden"
                      onChange={(e) => handleFileSelect(e, 'stalo')}
                    />
                  </div>
                ) : (
                  <div className="bg-[#eef2ec]/70 border border-[#cfd9d0] rounded-lg p-4 flex items-center justify-between">
                    <div className="flex items-center gap-3">
                      <div className="w-8 h-8 bg-[#1e7145] rounded flex items-center justify-center">
                        <svg className="w-5 h-5 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                        </svg>
                      </div>
                      <div>
                        <p className="font-mono text-sm font-semibold">{fileNameStalo}</p>
                        <p className="text-xs text-gray-500">Готов к обработке</p>
                      </div>
                    </div>
                    <button 
                      onClick={() => { setFileStalo(null); setFileNameStalo('') }}
                      className="btn-ghost text-sm"
                    >
                      Заменить
                    </button>
                  </div>
                )}
                
                <div className="flex gap-3 mt-6">
                  <button 
                    onClick={() => setStep(1)}
                    className="btn-ghost flex-1"
                  >
                    ← Назад
                  </button>
                  {fileStalo && (
                    <button 
                      onClick={compareFiles}
                      disabled={loading}
                      className="btn-primary flex-1"
                    >
                      {loading ? 'Обработка...' : 'Начать сравнение →'}
                    </button>
                  )}
                </div>
              </div>
            )}
            
            {/* Шаг 3: Обработка */}
            {step === 3 && (
              <div className="text-center py-12">
                <div className="spinner mx-auto mb-4"></div>
                <h2 className="font-unbounded text-lg font-semibold mb-2">Обработка данных</h2>
                <p className="text-gray-500">Сравниваем позиции, рассчитываем суммы...</p>
              </div>
            )}
            
            {/* Шаг 4: Результат */}
            {step === 4 && summary && (
              <div>
                <div className="flex items-center gap-2 mb-4">
                  <svg className="w-5 h-5 text-[#1e7145]" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z" />
                  </svg>
                  <h2 className="font-unbounded text-lg font-semibold">Результат сравнения</h2>
                </div>
                
                {/* Проверочные суммы */}
                <div className="bg-[#f0f7f2] border border-[#cfd9d0] rounded-lg p-4 mb-6">
                  <h3 className="font-semibold mb-3 text-sm">Проверочные суммы:</h3>
                  <div className="grid grid-cols-2 gap-4 text-sm">
                    <div>
                      <p className="text-gray-500 text-xs mb-1">БЫЛО:</p>
                      <p className="font-mono">СМР: {formatMoney(summary.smrBylo)}</p>
                      <p className="font-mono">ТМЦ: {formatMoney(summary.tmcBylo)}</p>
                      <p className="font-mono font-semibold">Всего: {formatMoney(summary.totalBylo)}</p>
                    </div>
                    <div>
                      <p className="text-gray-500 text-xs mb-1">СТАЛО:</p>
                      <p className="font-mono">СМР: {formatMoney(summary.smrStalo)}</p>
                      <p className="font-mono">ТМЦ: {formatMoney(summary.tmcStalo)}</p>
                      <p className="font-mono font-semibold">Всего: {formatMoney(summary.totalStalo)}</p>
                    </div>
                  </div>
                </div>
                
                {/* Кнопки экспорта */}
                <div className="flex gap-3">
                  <button onClick={handleReset} className="btn-ghost flex-1">
                    Начать заново
                  </button>
                  <button className="btn-primary flex-1">
                    Экспорт в Excel
                  </button>
                </div>
              </div>
            )}
          </div>
          
          {/* Правый сайдбар */}
          <div className="lg:sticky lg:top-[185px] lg:self-start space-y-4">
            {/* Предпросмотр */}
            <div className="card p-4">
              <div className="flex items-center justify-between mb-3">
                <h3 className="font-semibold text-sm">Структура таблицы</h3>
                <span className="text-xs bg-[#e0efe5] text-[#123f28] px-2 py-0.5 rounded">10 колонок</span>
              </div>
              <table className="mini-excel w-full">
                <thead>
                  <tr>
                    <th className="bg-[#D8D8D8]">A</th>
                    <th className="bg-[#D8D8D8]">B</th>
                    <th className="bg-[#D8D8D8]">C</th>
                    <th className="bg-[#D8D8D8]">D</th>
                    <th className="bg-[#D8D8D8]">E</th>
                    <th className="bg-[#D8D8D8]">F</th>
                    <th className="bg-[#1e7145]/20">G</th>
                    <th className="bg-[#1e7145]/20">H</th>
                    <th className="bg-[#1e7145]/20">I</th>
                    <th className="bg-[#dc2626]/20">J</th>
                  </tr>
                </thead>
                <tbody>
                  <tr>
                    <td className="cell-key">◆</td>
                    <td>Наименование</td>
                    <td>Ед.</td>
                    <td>Объём</td>
                    <td>Цена</td>
                    <td>Сумма</td>
                    <td className="cell-new">Объём</td>
                    <td className="cell-new">Цена</td>
                    <td className="cell-new">Сумма</td>
                    <td className="cell-new">Δ</td>
                  </tr>
                </tbody>
              </table>
              <div className="mt-3 text-xs text-gray-500 flex gap-3">
                <span className="flex items-center gap-1">
                  <span className="w-3 h-3 bg-[#D8D8D8] rounded"></span> Шапка
                </span>
                <span className="flex items-center gap-1">
                  <span className="w-3 h-3 bg-[#1e7145]/20 rounded"></span> СТАЛО
                </span>
                <span className="flex items-center gap-1">
                  <span className="w-3 h-3 bg-[#dc2626]/20 rounded"></span> Разница
                </span>
              </div>
            </div>
            
            {/* Регламент */}
            <div className="card p-4 bg-[#f0f7f2]/60 border-[#1e7145]">
              <div className="flex items-center gap-2 mb-3">
                <svg className="w-5 h-5 text-[#1e7145]" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                </svg>
                <h3 className="font-semibold text-sm">Как это работает</h3>
              </div>
              <ol className="text-xs space-y-2 text-gray-600">
                <li className="flex gap-2">
                  <span className="font-bold text-[#1e7145]">1.</span>
                  Загрузите файл «БЫЛО» с листами «СМР уник» и «ТМЦ уник»
                </li>
                <li className="flex gap-2">
                  <span className="font-bold text-[#1e7145]">2.</span>
                  Загрузите файл «СТАЛО» с аналогичной структурой
                </li>
                <li className="flex gap-2">
                  <span className="font-bold text-[#1e7145]">3.</span>
                  Система сопоставит позиции по кодам (ИД.КЕР/ТМЦ)
                </li>
                <li className="flex gap-2">
                  <span className="font-bold text-[#1e7145]">4.</span>
                  Рассчитаются суммы и разница (СТАЛО − БЫЛО)
                </li>
                <li className="flex gap-2">
                  <span className="font-bold text-[#1e7145]">5.</span>
                  Новые позиции добавятся автоматически
                </li>
                <li className="flex gap-2">
                  <span className="font-bold text-[#1e7145]">6.</span>
                  Получите готовую таблицу с проверочными суммами
                </li>
              </ol>
            </div>
            
            {/* О приложении */}
            <div className="card p-4">
              <h3 className="font-semibold text-sm mb-2">О приложении</h3>
              <p className="text-xs text-gray-500 mb-2">
                StateDiff — веб-аналог VBA макроса для сравнения сметных данных.
              </p>
              <p className="text-xs text-gray-500">
                Все вычисления выполняются локально в вашем браузере. Данные не отправляются на сервер.
              </p>
            </div>
          </div>
        </div>
      </main>
    </div>
  )
}

export default App
