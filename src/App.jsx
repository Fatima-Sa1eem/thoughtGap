import { useState, useEffect, useRef, useCallback, useMemo } from 'react'
import { Document, Packer, Paragraph, TextRun } from 'docx'

const PRESETS = {
  warmSepia: { bg: '#f4ecd8', text: '#5c4b37', name: 'Warm Sepia' },
  oledDark: { bg: '#000000', text: '#e0e0e0', name: 'OLED Dark' },
  softForest: { bg: '#1a2f1a', text: '#d4e5d4', name: 'Soft Forest' },
  creamPaper: { bg: '#faf8f0', text: '#2c2c2c', name: 'Cream/Paper' },
}

const FONTS = {
  monospace: { family: 'ui-monospace, monospace', name: 'Monospace' },
  serif: { family: 'Merriweather, serif', name: 'Serif (Merriweather)' },
  sans: { family: 'Inter, sans-serif', name: 'Sans-Serif (Inter)' },
}

const DIVIDER_LINE = '\n\n_____________\n\n'

function App() {
  const [text, setText] = useState('')
  const [showSettings, setShowSettings] = useState(false)
  const [showTimer, setShowTimer] = useState(false)
  const [timerMinutes, setTimerMinutes] = useState(15)
  const [timerActive, setTimerActive] = useState(false)
  const [timerSeconds, setTimerSeconds] = useState(0)
  const [locked, setLocked] = useState(false)
  const [timerMode, setTimerMode] = useState('time') // 'time', 'words', or 'session'
  const [wordGoal, setWordGoal] = useState(500)
  const [sessionMode, setSessionMode] = useState(false)
  const [sessionPhase, setSessionPhase] = useState('brainstorm') // 'brainstorm', 'write', 'edit'
  const [sessionTimers, setSessionTimers] = useState({ brainstorm: 10, write: 30, edit: 10 })
  const [sessionSeconds, setSessionSeconds] = useState(0)
  const [sessions, setSessions] = useState([])
  const [currentSessionId, setCurrentSessionId] = useState(null)
  const [brainstormNotes, setBrainstormNotes] = useState('')
  const [writingPaused, setWritingPaused] = useState(false)
  const [lastWritingKeystroke, setLastWritingKeystroke] = useState(Date.now())
  const [showSessions, setShowSessions] = useState(false)
  const [customBg, setCustomBg] = useState('#faf8f0')
  const [customText, setCustomText] = useState('#2c2c2c')
  const [selectedPreset, setSelectedPreset] = useState('creamPaper')
  const [selectedFont, setSelectedFont] = useState('sans')
  const [typewriterMode, setTypewriterMode] = useState(false)
  const [thoughtBreakMinutes, setThoughtBreakMinutes] = useState(8)
  const [showBackspaceWarning, setShowBackspaceWarning] = useState(false)
  const [lastKeystroke, setLastKeystroke] = useState(Date.now())
  const [dividerJustInserted, setDividerJustInserted] = useState(false)
  const [isFullscreen, setIsFullscreen] = useState(false)

  const textareaRef = useRef(null)
  const timerRef = useRef(null)
  const inactivityRef = useRef(null)

  const wordCount = useMemo(() => {
    return text.trim() ? text.trim().split(/\s+/).length : 0
  }, [text])

  const insertDivider = useCallback(() => {
    const textarea = textareaRef.current
    if (!textarea) return

    const start = textarea.selectionStart
    const end = textarea.selectionEnd
    
    setText(prevText => {
      const newText = prevText.slice(0, start) + DIVIDER_LINE + prevText.slice(end)
      setDividerJustInserted(true)
      
      setTimeout(() => {
        textarea.selectionStart = textarea.selectionEnd = start + DIVIDER_LINE.length
      }, 0)
      
      return newText
    })
  }, [])

  const toggleFullscreen = useCallback(() => {
    if (document.fullscreenElement) {
      document.exitFullscreen()
      setIsFullscreen(false)
    } else {
      document.documentElement.requestFullscreen()
      setIsFullscreen(true)
    }
  }, [])

  // Load from localStorage on mount
  useEffect(() => {
    const saved = localStorage.getItem('thoughtgap-content')
    const savedSettings = localStorage.getItem('thoughtgap-settings')
    const savedSessions = localStorage.getItem('thoughtgap-sessions')
    
    if (saved) setText(saved)
    
    if (savedSettings) {
      const settings = JSON.parse(savedSettings)
      setCustomBg(settings.customBg || '#faf8f0')
      setCustomText(settings.customText || '#2c2c2c')
      setSelectedPreset(settings.selectedPreset || 'creamPaper')
      setSelectedFont(settings.selectedFont || 'sans')
      setTypewriterMode(settings.typewriterMode || false)
      setThoughtBreakMinutes(settings.thoughtBreakMinutes || 8)
      setSessionTimers(settings.sessionTimers || { brainstorm: 10, write: 30, edit: 10 })
    }

    if (savedSessions) {
      setSessions(JSON.parse(savedSessions))
    }
  }, [])

  // Auto-save to localStorage
  useEffect(() => {
    localStorage.setItem('thoughtgap-content', text)
    localStorage.setItem('thoughtgap-settings', JSON.stringify({
      customBg,
      customText,
      selectedPreset,
      selectedFont,
      typewriterMode,
      thoughtBreakMinutes,
      sessionTimers,
    }))
    localStorage.setItem('thoughtgap-sessions', JSON.stringify(sessions))
  }, [text, customBg, customText, selectedPreset, selectedFont, typewriterMode, thoughtBreakMinutes, sessionTimers, sessions])

  // Inactivity tracking for thought break
  useEffect(() => {
    const resetInactivity = () => {
      setLastKeystroke(Date.now())
      setDividerJustInserted(false)
    }

    const checkInactivity = () => {
      const now = Date.now()
      const elapsed = (now - lastKeystroke) / 1000 / 60 // minutes
      
      if (elapsed >= thoughtBreakMinutes && !dividerJustInserted && !locked) {
        insertDivider()
      }
    }

    const handleGlobalKeyDown = (e) => {
      resetInactivity()
      // F key to toggle fullscreen (only when not in textarea or with modifier)
      if (e.key === 'f' && !e.ctrlKey && !e.metaKey && !e.altKey && document.activeElement !== textareaRef.current) {
        e.preventDefault()
        toggleFullscreen()
      }
    }

    inactivityRef.current = setInterval(checkInactivity, 1000)

    window.addEventListener('keydown', handleGlobalKeyDown)
    window.addEventListener('click', resetInactivity)

    return () => {
      clearInterval(inactivityRef.current)
      window.removeEventListener('keydown', handleGlobalKeyDown)
      window.removeEventListener('click', resetInactivity)
    }
  }, [lastKeystroke, thoughtBreakMinutes, dividerJustInserted, locked, insertDivider, toggleFullscreen])

  // Timer countdown and word count tracking
  useEffect(() => {
    if (timerActive) {
      if (timerMode === 'time') {
        if (timerSeconds > 0) {
          timerRef.current = setInterval(() => {
            setTimerSeconds(prev => prev - 1)
          }, 1000)
        } else if (timerSeconds === 0) {
          setTimerActive(false)
          setLocked(false)
          if (document.fullscreenElement) {
            document.exitFullscreen()
          }
        }
      } else if (timerMode === 'words') {
        // Word count mode - check if goal reached
        if (wordCount >= wordGoal) {
          setTimerActive(false)
          setLocked(false)
          if (document.fullscreenElement) {
            document.exitFullscreen()
          }
        }
      } else if (timerMode === 'session' && sessionMode) {
        // Session mode timer
        if (sessionSeconds > 0) {
          timerRef.current = setInterval(() => {
            setSessionSeconds(prev => prev - 1)
          }, 1000)
        } else if (sessionSeconds === 0) {
          // Phase complete, move to next phase
          if (sessionPhase === 'brainstorm') {
            setSessionPhase('write')
            setSessionSeconds(sessionTimers.write * 60)
            setWritingPaused(false)
          } else if (sessionPhase === 'write') {
            // Lock writing and move to edit phase
            setSessionPhase('edit')
            setSessionSeconds(sessionTimers.edit * 60)
            // Create session copy
            const newSession = {
              id: Date.now(),
              brainstorm: brainstormNotes,
              writing: text,
              timestamp: new Date().toISOString(),
            }
            setSessions(prev => [...prev, newSession])
            setCurrentSessionId(newSession.id)
          } else if (sessionPhase === 'edit') {
            // Session complete
            setSessionMode(false)
            setTimerActive(false)
            setLocked(false)
            if (document.fullscreenElement) {
              document.exitFullscreen()
            }
          }
        }
      }
    }

    return () => clearInterval(timerRef.current)
  }, [timerActive, timerSeconds, timerMode, wordCount, wordGoal, sessionMode, sessionPhase, sessionSeconds, sessionTimers, brainstormNotes, text])

  // Writing phase inactivity detection (50 seconds)
  useEffect(() => {
    if (sessionMode && sessionPhase === 'write' && timerActive && !writingPaused) {
      const checkWritingInactivity = () => {
        const now = Date.now()
        const elapsed = (now - lastWritingKeystroke) / 1000
        if (elapsed >= 50) {
          setWritingPaused(true)
        }
      }

      const interval = setInterval(checkWritingInactivity, 1000)
      return () => clearInterval(interval)
    }
  }, [sessionMode, sessionPhase, timerActive, writingPaused, lastWritingKeystroke])

  const handleKeyDown = (e) => {
    // Escape to toggle settings
    if (e.key === 'Escape') {
      if (locked) {
        // Allow early exit from focus mode
        setTimerActive(false)
        setLocked(false)
        setTimerSeconds(0)
        if (document.fullscreenElement) {
          document.exitFullscreen()
        }
      } else {
        setShowSettings(!showSettings)
      }
    }

    // Backspace after divider
    if (e.key === 'Backspace' && dividerJustInserted) {
      const textarea = textareaRef.current
      const cursorPos = textarea.selectionStart
      const textBeforeCursor = text.slice(0, cursorPos)
      
      if (textBeforeCursor.endsWith(DIVIDER_LINE)) {
        e.preventDefault()
        const newText = text.slice(0, cursorPos - DIVIDER_LINE.length) + text.slice(cursorPos)
        setText(newText)
        setDividerJustInserted(false)
        textarea.selectionStart = textarea.selectionEnd = cursorPos - DIVIDER_LINE.length
      }
    }

    // Prevent backspace in typewriter mode
    if (e.key === 'Backspace' && typewriterMode && !dividerJustInserted) {
      e.preventDefault()
      setShowBackspaceWarning(true)
      setTimeout(() => setShowBackspaceWarning(false), 2000)
    }

    // Track writing keystrokes in session mode
    if (sessionMode && sessionPhase === 'write' && !writingPaused) {
      setLastWritingKeystroke(Date.now())
    }
  }

  const startFocusTimer = () => {
    if (timerMode === 'session') {
      // Save current content as a session before starting new one
      if (text.trim() || brainstormNotes.trim()) {
        const newSession = {
          id: Date.now(),
          brainstorm: brainstormNotes,
          writing: text,
          timestamp: new Date().toISOString(),
        }
        setSessions(prev => [...prev, newSession])
      }
      // Clear for new session
      setText('')
      setBrainstormNotes('')
      setCurrentSessionId(null)
      
      setSessionMode(true)
      setSessionPhase('brainstorm')
      setSessionSeconds(sessionTimers.brainstorm * 60)
      setWritingPaused(false)
      setLastWritingKeystroke(Date.now())
    } else if (timerMode === 'time') {
      const totalSeconds = timerMinutes * 60
      setTimerSeconds(totalSeconds)
    }
    setTimerActive(true)
    setLocked(true)
    setShowTimer(false)
    
    // Request fullscreen
    if (document.documentElement.requestFullscreen) {
      document.documentElement.requestFullscreen()
    }
  }

  const resumeWriting = () => {
    setWritingPaused(false)
    setLastWritingKeystroke(Date.now())
  }

  const loadSession = (sessionId) => {
    const session = sessions.find(s => s.id === sessionId)
    if (session) {
      setText(session.writing)
      setBrainstormNotes(session.brainstorm)
      setCurrentSessionId(sessionId)
      setShowSessions(false)
    }
  }

  const deleteSession = (sessionId) => {
    setSessions(prev => prev.filter(s => s.id !== sessionId))
  }

  const formatSessionTime = (seconds) => {
    const mins = Math.floor(seconds / 60)
    const secs = seconds % 60
    return `${mins}:${secs.toString().padStart(2, '0')}`
  }

  // Prevent tab closure during timer
  useEffect(() => {
    const handleBeforeUnload = (e) => {
      if (timerActive) {
        e.preventDefault()
        e.returnValue = ''
      }
    }

    window.addEventListener('beforeunload', handleBeforeUnload)
    return () => window.removeEventListener('beforeunload', handleBeforeUnload)
  }, [timerActive])

  // Sync fullscreen state with browser
  useEffect(() => {
    const handleFullscreenChange = () => {
      setIsFullscreen(!!document.fullscreenElement)
    }

    document.addEventListener('fullscreenchange', handleFullscreenChange)
    return () => document.removeEventListener('fullscreenchange', handleFullscreenChange)
  }, [])

  const exportAsTxt = () => {
    const blob = new Blob([text], { type: 'text/plain' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = 'document.txt'
    a.click()
    URL.revokeObjectURL(url)
  }

  const exportAsDocx = async () => {
    const doc = new Document({
      sections: [{
        properties: {},
        children: text.split('\n').map(line => 
          new Paragraph({
            children: [new TextRun(line)]
          })
        )
      }]
    })

    const blob = await Packer.toBlob(doc)
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = 'document.docx'
    a.click()
    URL.revokeObjectURL(url)
  }

  const handleFileImport = (e) => {
    const file = e.target.files[0]
    if (!file) return

    const reader = new FileReader()
    reader.onload = (event) => {
      setText(event.target.result)
    }
    reader.readAsText(file)
  }

  const handleDrop = (e) => {
    e.preventDefault()
    const file = e.dataTransfer.files[0]
    if (!file) return

    const reader = new FileReader()
    reader.onload = (event) => {
      setText(event.target.result)
    }
    reader.readAsText(file)
  }

  const handleDragOver = (e) => {
    e.preventDefault()
  }

  const applyPreset = (presetKey) => {
    const preset = PRESETS[presetKey]
    setCustomBg(preset.bg)
    setCustomText(preset.text)
    setSelectedPreset(presetKey)
  }

  const readingTime = Math.ceil(wordCount / 200) // Average 200 words per minute

  // Typewriter mode - scroll to keep cursor centered
  const handleInput = useCallback(() => {
    if (typewriterMode && textareaRef.current) {
      const textarea = textareaRef.current
      const lineHeight = 24 // Approximate line height
      const cursorPosition = textarea.selectionStart
      const textBeforeCursor = text.substring(0, cursorPosition)
      const lines = textBeforeCursor.split('\n')
      const currentLine = lines.length
      
      const scrollTop = (currentLine * lineHeight) - (textarea.clientHeight / 2)
      textarea.scrollTop = Math.max(0, scrollTop)
    }
  }, [typewriterMode, text])

  const formatTime = (seconds) => {
    const mins = Math.floor(seconds / 60)
    const secs = seconds % 60
    return `${mins}:${secs.toString().padStart(2, '0')}`
  }

  const currentBg = PRESETS[selectedPreset]?.bg || customBg
  const currentText = PRESETS[selectedPreset]?.text || customText
  const currentFont = FONTS[selectedFont]?.family || 'Inter, sans-serif'

  return (
    <div 
      className="w-screen h-screen relative"
      style={{ backgroundColor: currentBg, fontFamily: currentFont }}
      onDrop={handleDrop}
      onDragOver={handleDragOver}
    >
      {/* Main Editor */}
      <textarea
        ref={textareaRef}
        value={text}
        onChange={(e) => setText(e.target.value)}
        onKeyDown={handleKeyDown}
        onInput={handleInput}
        className="w-full h-full p-8 resize-none outline-none bg-transparent"
        style={{ 
          color: currentText, 
          fontSize: '18px',
          lineHeight: '1.8',
          fontFamily: currentFont,
        }}
        placeholder="Start writing..."
      />

      {/* Settings Toggle */}
      {!locked && (
        <button
          onClick={() => setShowSettings(!showSettings)}
          className="absolute top-4 right-4 p-2 rounded opacity-30 hover:opacity-100 transition-opacity duration-200"
          style={{ color: currentText, backgroundColor: `${currentText}20` }}
          title="Settings (Esc)"
        >
          <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <circle cx="12" cy="12" r="3"></circle>
            <path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z"></path>
          </svg>
        </button>
      )}

      {/* Sessions Toggle */}
      {!locked && (
        <button
          onClick={() => setShowSessions(!showSessions)}
          className="absolute top-4 right-28 p-2 rounded opacity-30 hover:opacity-100 transition-opacity duration-200"
          style={{ color: currentText, backgroundColor: `${currentText}20` }}
          title="Sessions"
        >
          <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
            <polyline points="14 2 14 8 20 8"></polyline>
            <line x1="16" y1="13" x2="8" y2="13"></line>
            <line x1="16" y1="17" x2="8" y2="17"></line>
            <polyline points="10 9 9 9 8 9"></polyline>
          </svg>
        </button>
      )}

      {/* Fullscreen Toggle */}
      {!locked && (
        <button
          onClick={toggleFullscreen}
          className="absolute top-4 right-16 p-2 rounded opacity-30 hover:opacity-100 transition-opacity duration-200"
          style={{ color: currentText, backgroundColor: `${currentText}20` }}
          title="Fullscreen (F)"
        >
          {isFullscreen ? (
            <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <path d="M8 3v3a2 2 0 0 1-2 2H3"></path>
              <path d="M21 8v-3a2 2 0 0 0-2-2h-3"></path>
              <path d="M3 16v3a2 2 0 0 0 2 2h3"></path>
              <path d="M16 21h3a2 2 0 0 0 2-2v-3"></path>
            </svg>
          ) : (
            <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <path d="M8 3H5a2 2 0 0 0-2 2v3"></path>
              <path d="M21 8V5a2 2 0 0 0-2-2h-3"></path>
              <path d="M3 16v3a2 2 0 0 0 2 2h3"></path>
              <path d="M16 21h3a2 2 0 0 0 2-2v-3"></path>
            </svg>
          )}
        </button>
      )}

      {/* Sessions Panel */}
      {showSessions && !locked && (
        <div 
          className="absolute top-16 right-4 p-4 rounded-lg shadow-xl w-72 max-h-[70vh] overflow-y-auto"
          style={{ backgroundColor: currentBg, color: currentText, border: `1px solid ${currentText}30` }}
        >
          <div className="flex justify-between items-center mb-4">
            <h2 className="text-sm font-semibold">Writing Sessions</h2>
            <button
              onClick={() => {
                // Save current content as a session if there's any content
                if (text.trim() || brainstormNotes.trim()) {
                  const newSession = {
                    id: Date.now(),
                    brainstorm: brainstormNotes,
                    writing: text,
                    timestamp: new Date().toISOString(),
                  }
                  setSessions(prev => [...prev, newSession])
                }
                // Clear current content for new session
                setText('')
                setBrainstormNotes('')
                setCurrentSessionId(null)
                setTimerMode('session')
                setShowSettings(true)
                setShowSessions(false)
              }}
              className="p-1 rounded"
              style={{ backgroundColor: `${currentText}20`, color: currentText }}
              title="Create new session"
            >
              <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <line x1="12" y1="5" x2="12" y2="19"></line>
                <line x1="5" y1="12" x2="19" y2="12"></line>
              </svg>
            </button>
          </div>

          {/* Current Session */}
          {sessionMode && (
            <div 
              className="p-3 rounded mb-3 border"
              style={{ borderColor: `${currentText}40`, backgroundColor: `${currentText}10` }}
            >
              <div className="text-xs font-semibold mb-1">📝 Current Session</div>
              <div className="text-xs opacity-70 capitalize mb-1">{sessionPhase} Phase</div>
              <div className="text-xs opacity-50">{formatSessionTime(sessionSeconds)} remaining</div>
            </div>
          )}

          {/* Past Sessions */}
          <div className="space-y-2">
            {sessions.length === 0 ? (
              <p className="text-xs opacity-50 text-center py-4">No sessions yet. Click + to start.</p>
            ) : (
              sessions.map(session => (
                <div 
                  key={session.id}
                  className="p-3 rounded border cursor-pointer hover:opacity-80 transition-opacity"
                  style={{ borderColor: `${currentText}20`, backgroundColor: `${currentText}5` }}
                  onClick={() => loadSession(session.id)}
                >
                  <div className="flex justify-between items-start mb-1">
                    <div className="text-xs opacity-70">
                      {new Date(session.timestamp).toLocaleDateString()}
                    </div>
                    <button
                      onClick={(e) => {
                        e.stopPropagation()
                        deleteSession(session.id)
                      }}
                      className="text-xs opacity-50 hover:opacity-100"
                      style={{ color: currentText }}
                    >
                      ✕
                    </button>
                  </div>
                  <div className="text-xs opacity-50 truncate">
                    {session.writing.slice(0, 40)}...
                  </div>
                </div>
              ))
            )}
          </div>
        </div>
      )}

      {/* Settings Panel */}
      {showSettings && !locked && (
        <div 
          className="absolute top-16 right-4 p-6 rounded-lg shadow-xl w-80"
          style={{ backgroundColor: currentBg, color: currentText, border: `1px solid ${currentText}30` }}
        >
          <h2 className="text-lg font-semibold mb-4">Settings</h2>
          
          {/* Color Presets */}
          <div className="mb-4">
            <label className="block text-sm mb-2 opacity-70">Color Presets</label>
            <div className="grid grid-cols-2 gap-2">
              {Object.entries(PRESETS).map(([key, preset]) => (
                <button
                  key={key}
                  onClick={() => applyPreset(key)}
                  className={`p-2 rounded text-xs transition-all ${
                    selectedPreset === key ? 'ring-2 ring-offset-2' : 'opacity-70 hover:opacity-100'
                  }`}
                  style={{ 
                    backgroundColor: preset.bg, 
                    color: preset.text,
                    borderColor: currentText
                  }}
                >
                  {preset.name}
                </button>
              ))}
            </div>
          </div>

          {/* Custom Colors */}
          <div className="mb-4 grid grid-cols-2 gap-4">
            <div>
              <label className="block text-sm mb-2 opacity-70">Background</label>
              <input
                type="color"
                value={customBg}
                onChange={(e) => {
                  setCustomBg(e.target.value)
                  setSelectedPreset(null)
                }}
                className="w-full h-8 rounded cursor-pointer"
              />
            </div>
            <div>
              <label className="block text-sm mb-2 opacity-70">Text</label>
              <input
                type="color"
                value={customText}
                onChange={(e) => {
                  setCustomText(e.target.value)
                  setSelectedPreset(null)
                }}
                className="w-full h-8 rounded cursor-pointer"
              />
            </div>
          </div>

          {/* Font Selection */}
          <div className="mb-4">
            <label className="block text-sm mb-2 opacity-70">Font</label>
            <select
              value={selectedFont}
              onChange={(e) => setSelectedFont(e.target.value)}
              className="w-full p-2 rounded bg-transparent border"
              style={{ borderColor: `${currentText}30`, color: currentText }}
            >
              {Object.entries(FONTS).map(([key, font]) => (
                <option key={key} value={key} style={{ color: '#000' }}>
                  {font.name}
                </option>
              ))}
            </select>
          </div>

          {/* Typewriter Mode */}
          <div className="mb-4">
            <div className="flex items-center justify-between mb-1">
              <label className="text-sm opacity-70">Typewriter Mode</label>
              <button
                onClick={() => setTypewriterMode(!typewriterMode)}
                className={`w-12 h-6 rounded-full transition-colors ${
                  typewriterMode ? 'opacity-100' : 'opacity-50'
                }`}
                style={{ backgroundColor: typewriterMode ? currentText : `${currentText}30` }}
              >
                <div
                  className={`w-5 h-5 rounded-full transition-transform ${
                    typewriterMode ? 'translate-x-6' : 'translate-x-0.5'
                  }`}
                  style={{ backgroundColor: currentBg }}
                />
              </button>
            </div>
            <p className="text-xs opacity-50">Disables backspace for forward-only writing</p>
          </div>

          {/* Thought Break Timer */}
          <div className="mb-4">
            <label className="block text-sm mb-2 opacity-70">Thought Break (minutes)</label>
            <input
              type="number"
              value={thoughtBreakMinutes}
              onChange={(e) => setThoughtBreakMinutes(Math.max(1, parseInt(e.target.value) || 8))}
              className="w-full p-2 rounded bg-transparent border"
              style={{ borderColor: `${currentText}30`, color: currentText }}
              min="1"
            />
          </div>

          {/* Export/Import */}
          <div className="mb-4 grid grid-cols-2 gap-2">
            <button
              onClick={exportAsTxt}
              className="p-2 rounded text-xs hover:opacity-80 transition-opacity"
              style={{ backgroundColor: `${currentText}20`, color: currentText }}
            >
              Export .txt
            </button>
            <button
              onClick={exportAsDocx}
              className="p-2 rounded text-xs hover:opacity-80 transition-opacity"
              style={{ backgroundColor: `${currentText}20`, color: currentText }}
            >
              Export .docx
            </button>
          </div>

          <div className="mb-4">
            <label className="block text-sm mb-2 opacity-70">Import File</label>
            <input
              type="file"
              accept=".txt,.docx"
              onChange={handleFileImport}
              className="w-full text-xs"
            />
          </div>

          {/* Focus Timer */}
          <div className="border-t pt-4" style={{ borderColor: `${currentText}20` }}>
            <label className="block text-sm mb-2 opacity-70">Focus Timer</label>
            
            {/* Timer Mode Selection */}
            <div className="flex gap-2 mb-3">
              <button
                onClick={() => setTimerMode('time')}
                className={`flex-1 p-2 rounded text-xs ${
                  timerMode === 'time' ? 'opacity-100' : 'opacity-50 hover:opacity-80'
                }`}
                style={{ backgroundColor: timerMode === 'time' ? currentText : `${currentText}20`, color: timerMode === 'time' ? currentBg : currentText }}
              >
                Time
              </button>
              <button
                onClick={() => setTimerMode('words')}
                className={`flex-1 p-2 rounded text-xs ${
                  timerMode === 'words' ? 'opacity-100' : 'opacity-50 hover:opacity-80'
                }`}
                style={{ backgroundColor: timerMode === 'words' ? currentText : `${currentText}20`, color: timerMode === 'words' ? currentBg : currentText }}
              >
                Words
              </button>
              <button
                onClick={() => setTimerMode('session')}
                className={`flex-1 p-2 rounded text-xs ${
                  timerMode === 'session' ? 'opacity-100' : 'opacity-50 hover:opacity-80'
                }`}
                style={{ backgroundColor: timerMode === 'session' ? currentText : `${currentText}20`, color: timerMode === 'session' ? currentBg : currentText }}
              >
                10-30-10
              </button>
            </div>

            {timerMode === 'time' && (
              <div className="flex gap-2 mb-2">
                {[15, 30, 45].map((mins) => (
                  <button
                    key={mins}
                    onClick={() => {
                      setTimerMinutes(mins)
                      setShowTimer(true)
                    }}
                    className={`flex-1 p-2 rounded text-xs ${
                      timerMinutes === mins ? 'opacity-100' : 'opacity-50 hover:opacity-80'
                    }`}
                    style={{ backgroundColor: `${currentText}20`, color: currentText }}
                  >
                    {mins}m
                  </button>
                ))}
              </div>
            )}

            {timerMode === 'words' && (
              <div className="mb-2">
                <input
                  type="number"
                  value={wordGoal}
                  onChange={(e) => setWordGoal(Math.max(1, parseInt(e.target.value) || 500))}
                  className="w-full p-2 rounded bg-transparent border text-center"
                  style={{ borderColor: `${currentText}30`, color: currentText }}
                  min="1"
                />
                <p className="text-xs opacity-50 text-center mt-1">Word goal</p>
              </div>
            )}

            {timerMode === 'session' && (
              <div className="mb-2 grid grid-cols-3 gap-2">
                <div>
                  <label className="text-xs opacity-50 block mb-1">Brainstorm</label>
                  <input
                    type="number"
                    value={sessionTimers.brainstorm}
                    onChange={(e) => setSessionTimers(prev => ({ ...prev, brainstorm: Math.max(1, parseInt(e.target.value) || 10) }))}
                    className="w-full p-2 rounded bg-transparent border text-center"
                    style={{ borderColor: `${currentText}30`, color: currentText }}
                    min="1"
                  />
                </div>
                <div>
                  <label className="text-xs opacity-50 block mb-1">Write</label>
                  <input
                    type="number"
                    value={sessionTimers.write}
                    onChange={(e) => setSessionTimers(prev => ({ ...prev, write: Math.max(1, parseInt(e.target.value) || 30) }))}
                    className="w-full p-2 rounded bg-transparent border text-center"
                    style={{ borderColor: `${currentText}30`, color: currentText }}
                    min="1"
                  />
                </div>
                <div>
                  <label className="text-xs opacity-50 block mb-1">Edit</label>
                  <input
                    type="number"
                    value={sessionTimers.edit}
                    onChange={(e) => setSessionTimers(prev => ({ ...prev, edit: Math.max(1, parseInt(e.target.value) || 10) }))}
                    className="w-full p-2 rounded bg-transparent border text-center"
                    style={{ borderColor: `${currentText}30`, color: currentText }}
                    min="1"
                  />
                </div>
              </div>
            )}

            <button
              onClick={startFocusTimer}
              className="w-full p-2 rounded text-xs font-semibold"
              style={{ backgroundColor: currentText, color: currentBg }}
            >
              Start Focus Session
            </button>
          </div>
        </div>
      )}

      {/* Focus Timer Display */}
      {locked && (
        <div className="absolute top-4 left-1/2 transform -translate-x-1/2 text-center">
          <div 
            className="text-2xl font-mono px-4 py-2 rounded"
            style={{ color: currentText, backgroundColor: `${currentText}10` }}
          >
            {timerMode === 'time' ? formatTime(timerSeconds) : 
             timerMode === 'words' ? `${wordCount}/${wordGoal} words` :
             `${sessionPhase.toUpperCase()}: ${formatSessionTime(sessionSeconds)}`}
          </div>
          {timerMode === 'words' && (
            <div className="text-xs opacity-50 mt-1">
              {Math.round((wordCount / wordGoal) * 100)}% complete
            </div>
          )}
          {timerMode === 'session' && (
            <div className="text-xs opacity-50 mt-1 capitalize">
              {sessionPhase} phase
            </div>
          )}
        </div>
      )}

      {/* Session Mode Brainstorm Panel */}
      {sessionMode && sessionPhase === 'brainstorm' && (
        <div className="absolute top-16 left-4 right-4">
          <div 
            className="p-4 rounded-lg"
            style={{ backgroundColor: `${currentText}10`, color: currentText, border: `1px solid ${currentText}30` }}
          >
            <h3 className="text-sm font-semibold mb-2">🧠 Brainstorm Notes</h3>
            <textarea
              value={brainstormNotes}
              onChange={(e) => setBrainstormNotes(e.target.value)}
              className="w-full h-24 p-2 rounded bg-transparent resize-none outline-none text-sm"
              style={{ color: currentText, border: `1px solid ${currentText}20` }}
              placeholder="Jot down your ideas here..."
            />
          </div>
        </div>
      )}

      {/* Writing Paused Indicator */}
      {sessionMode && sessionPhase === 'write' && writingPaused && (
        <div className="absolute top-16 left-1/2 transform -translate-x-1/2">
          <div 
            className="p-4 rounded-lg text-center"
            style={{ backgroundColor: `${currentText}20`, color: currentText }}
          >
            <p className="text-sm mb-2">⏸️ Writing paused (no activity for 50s)</p>
            <button
              onClick={resumeWriting}
              className="px-4 py-2 rounded text-xs font-semibold"
              style={{ backgroundColor: currentText, color: currentBg }}
            >
              Resume Writing
            </button>
          </div>
        </div>
      )}

      {/* Backspace Warning */}
      {showBackspaceWarning && (
        <div className="absolute top-4 left-1/2 transform -translate-x-1/2">
          <div 
            className="text-sm px-4 py-2 rounded"
            style={{ color: currentText, backgroundColor: `${currentText}20` }}
          >
            Backspace disabled in Typewriter Mode
          </div>
        </div>
      )}

      {/* Word Count & Reading Time */}
      <div 
        className="absolute bottom-4 right-4 text-xs opacity-30"
        style={{ color: currentText }}
      >
        {wordCount} words · {readingTime} min read
      </div>
    </div>
  )
}

export default App
