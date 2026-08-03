import { useEffect, useRef, useState } from 'react'
import {
  AI_PROVIDERS,
  activePreset,
  isLocalCliProvider,
  type AiModelPreset,
  type AiProviderId,
  type AiSettings,
} from '@genoffice/ai-provider'
import type { Lang } from '@genoffice/i18n'
import { tAiSettings } from './ai-settings-strings'
import { useAiSettingsDialog, type AiSettingsHost } from './use-ai-settings-dialog'

export interface AiModelPickerProps {
  /** the settings the panel is currently sending requests with */
  settings: AiSettings
  lang: Lang
  host: AiSettingsHost
  /** the switch, persisted: the caller keeps its own copy in sync from here */
  onChange: (settings: AiSettings) => void
  /** the host app's own footer-button class, when it wants one */
  className?: string | undefined
}

/**
 * Composer-footer control for switching between configured backends.
 *
 * The settings dialog is where a model is *set up* — key, endpoint, model id.
 * Switching between two that are already set up is a different, far more
 * frequent action ("try that again on Codex"), and going through a modal for
 * it is three clicks too many. This lists what is ready to use and switches on
 * one click; the dialog stays one row away for everything else.
 *
 * Each provider slot holds one model, so a provider *is* a configured model
 * here — hence rows named for the provider with its model underneath.
 */
export function AiModelPicker({ settings, lang, host, onChange, className }: AiModelPickerProps) {
  const t = (key: Parameters<typeof tAiSettings>[1]) => tAiSettings(lang, key)
  usePickerStyles()

  const [rect, setRect] = useState<{ left: number; bottom: number } | null>(null)
  /**
   * Backends that need asking about rather than reading off the settings file:
   * a CLI has to be installed, Genspark has to be signed in. Cached for the
   * life of the panel — probing spawns a subprocess, and neither answer
   * changes without the user leaving the app.
   */
  const [reachable, setReachable] = useState<Partial<Record<AiProviderId, boolean>>>({})
  const probed = useRef(false)
  const wrapRef = useRef<HTMLDivElement>(null)
  const buttonRef = useRef<HTMLButtonElement>(null)
  const open = rect !== null

  const { open: openSettings, dialog } = useAiSettingsDialog(host, lang, onChange)

  const provider = settings.provider
  const meta = AI_PROVIDERS.find((p) => p.id === provider) ?? AI_PROVIDERS[0]!
  const modelOf = (id: AiProviderId): string => {
    const model = settings.providers[id]?.model ?? ''
    // a CLI with no model runs whatever its own config selects, which is the
    // right default for someone who already set that up
    return model || (isLocalCliProvider(id) ? t('cliDefaultModel') : '')
  }

  /**
   * Show a provider when it is ready to use — plus the active one regardless,
   * so the button never claims to be on a backend the list does not contain.
   */
  const usable = (id: AiProviderId): boolean => {
    if (id === provider) return true
    if (id === 'genspark' || isLocalCliProvider(id)) return reachable[id] === true
    return !!settings.providers[id]?.hasApiKey
  }
  const options = AI_PROVIDERS.filter((p) => usable(p.id))

  const openList = () => {
    const box = buttonRef.current?.getBoundingClientRect()
    // anchored upwards: the composer sits at the bottom of the panel
    if (box) setRect({ left: box.left, bottom: window.innerHeight - box.top + 6 })
    if (probed.current) return
    probed.current = true
    for (const option of AI_PROVIDERS) {
      if (isLocalCliProvider(option.id) && host.aiCliStatus) {
        void host
          .aiCliStatus(option.id)
          .then((status) =>
            setReachable((prev) => ({ ...prev, [option.id]: status.installed === true })),
          )
          .catch(() => {})
      }
    }
    if (host.gskStatus) {
      void host
        .gskStatus()
        .then((status) => setReachable((prev) => ({ ...prev, genspark: status.loggedIn })))
        .catch(() => {})
    }
  }

  const closeList = () => setRect(null)

  // close when the pointer goes down outside, and when anything scrolls: the
  // fixed popup would otherwise detach from the button it belongs to
  useEffect(() => {
    if (!open) return
    const onPointerDown = (e: PointerEvent) => {
      const target = e.target as Element | null
      if (wrapRef.current?.contains(target as Node)) return
      if (target?.closest?.('.gso-ai-model-list')) return
      closeList()
    }
    const onKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') closeList()
    }
    const onScroll = (e: Event) => {
      const target = e.target as Element | null
      if (target instanceof Element && target.closest('.gso-ai-model-list')) return
      closeList()
    }
    document.addEventListener('pointerdown', onPointerDown, true)
    document.addEventListener('keydown', onKeyDown)
    window.addEventListener('scroll', onScroll, true)
    window.addEventListener('resize', onScroll)
    return () => {
      document.removeEventListener('pointerdown', onPointerDown, true)
      document.removeEventListener('keydown', onKeyDown)
      window.removeEventListener('scroll', onScroll, true)
      window.removeEventListener('resize', onScroll)
    }
  }, [open])

  // Sent whole, exactly as the dialog saves: keys were blanked on the way to
  // this renderer, and an empty key means "keep the stored one" in main.
  const apply = (next: AiSettings) => {
    closeList()
    void host.setAiSettings(next).then(() => onChange(next))
  }

  const chooseProvider = (id: AiProviderId) => {
    if (id === provider) return closeList()
    apply({ ...settings, provider: id })
  }

  /** a preset carries a model as well, so it writes into that provider's slot */
  const choosePreset = (preset: AiModelPreset) => {
    const config = settings.providers[preset.provider]
    apply({
      ...settings,
      provider: preset.provider,
      providers: {
        ...settings.providers,
        [preset.provider]: { ...(config ?? { apiKey: '' }), model: preset.model },
      },
    })
  }

  const presets = settings.presets ?? []
  const active = activePreset(settings)

  // "CLI default" on the button would not say *which* backend is running, so
  // an unset model falls back to the provider's name there; the list has room
  // for both and shows them together.
  const current = active?.name ?? settings.providers[provider]?.model ?? ''
  return (
    <div className="gso-ai-model" ref={wrapRef}>
      <button
        type="button"
        ref={buttonRef}
        className={`gso-ai-model-btn${className ? ` ${className}` : ''}`}
        aria-haspopup="listbox"
        aria-expanded={open}
        // the button shows the model, which is what changes most often; the
        // provider disambiguates it in the tooltip and in the list
        title={`${t('switchModel')} — ${meta.label}${current ? ` · ${current}` : ''}`}
        onClick={() => (open ? closeList() : openList())}
      >
        <span className="gso-ai-model-name">{current || meta.label}</span>
        <svg width="10" height="10" viewBox="0 0 12 12" aria-hidden="true">
          <path
            d="M2.5 4.5 6 8l3.5-3.5"
            stroke="currentColor"
            strokeWidth="1.6"
            strokeLinecap="round"
            strokeLinejoin="round"
            fill="none"
          />
        </svg>
      </button>
      {rect && (
        <ul
          className="gso-ai-model-list"
          role="listbox"
          aria-label={t('switchModel')}
          style={{ left: rect.left, bottom: rect.bottom }}
        >
          {presets.length > 0 && (
            <li className="gso-ai-model-head" aria-hidden="true">
              {t('presets')}
            </li>
          )}
          {presets.map((preset) => (
            <li key={`preset:${preset.name}`}>
              <button
                type="button"
                role="option"
                aria-selected={preset === active}
                className={`gso-ai-model-item${preset === active ? ' active' : ''}`}
                onClick={() => choosePreset(preset)}
              >
                <span className="gso-ai-model-item-label">{preset.name}</span>
                <span className="gso-ai-model-item-model">
                  {`${AI_PROVIDERS.find((p) => p.id === preset.provider)?.label ?? preset.provider}\u2009·\u2009${
                    preset.model || t('cliDefaultModel')
                  }`}
                </span>
              </button>
            </li>
          ))}
          {presets.length > 0 && (
            <li className="gso-ai-model-head" aria-hidden="true">
              {t('provider')}
            </li>
          )}
          {options.map((option) => (
            <li key={option.id}>
              <button
                type="button"
                role="option"
                aria-selected={option.id === provider && !active}
                className={`gso-ai-model-item${option.id === provider && !active ? ' active' : ''}`}
                onClick={() => chooseProvider(option.id)}
              >
                <span className="gso-ai-model-item-label">{option.label}</span>
                <span className="gso-ai-model-item-model">{modelOf(option.id)}</span>
              </button>
            </li>
          ))}
          <li>
            <button
              type="button"
              className="gso-ai-model-item gso-ai-model-more"
              onClick={() => {
                closeList()
                openSettings()
              }}
            >
              {t('moreModels')}
            </button>
          </li>
        </ul>
      )}
      {dialog}
    </div>
  )
}

// ── styles ───────────────────────────────────────────────
// Injected by the component, like the settings dialog: four apps mount this in
// four separate stylesheets, and four copies of the same rules would drift.

const STYLE_ID = 'gso-ai-model-picker-styles'

const STYLES = `
.gso-ai-model { position: relative; display: inline-flex; min-width: 0; }
.gso-ai-model-btn {
  display: inline-flex; align-items: center; gap: 4px; max-width: 190px;
  padding: 3px 7px; border: 1px solid #e2e5ea; border-radius: 999px;
  background: #fff; color: #4b5563; cursor: pointer;
  font: inherit; font-size: 12px; line-height: 18px;
}
.gso-ai-model-btn:hover { background: #f2f4f7; color: #1b1f24; }
.gso-ai-model-name { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.gso-ai-model-list {
  position: fixed; z-index: 9001; min-width: 200px; max-width: 300px;
  max-height: 320px; overflow-y: auto;
  margin: 0; padding: 4px; list-style: none;
  background: #fff; border: 1px solid #d3d7de; border-radius: 8px;
  box-shadow: 0 8px 24px rgba(15, 20, 30, 0.16);
  font: 13px/1.5 system-ui, -apple-system, "Segoe UI", sans-serif; color: #1b1f24;
}
.gso-ai-model-item {
  display: block; width: 100%; padding: 6px 9px;
  border: 0; border-radius: 5px; background: none; cursor: pointer;
  font: inherit; text-align: start; color: #374151;
}
.gso-ai-model-item:hover { background: #f2f4f7; }
.gso-ai-model-item.active { background: #eaf1fe; color: #1a54c9; }
.gso-ai-model-item-label { display: block; font-weight: 600; }
.gso-ai-model-item-model {
  display: block; font-size: 12px; color: #6b7280;
  overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
}
.gso-ai-model-item.active .gso-ai-model-item-model { color: #4b7bd4; }
.gso-ai-model-head {
  padding: 6px 9px 2px; color: #9aa1ac; font-size: 11px;
  font-weight: 600; letter-spacing: 0.04em; text-transform: uppercase;
}
.gso-ai-model-more {
  margin-top: 4px; border-top: 1px solid #eceef2; border-radius: 0 0 5px 5px;
  padding-top: 8px; color: #6b7280; font-size: 12px;
}
`

function usePickerStyles(): void {
  useEffect(() => {
    if (document.getElementById(STYLE_ID)) return
    const style = document.createElement('style')
    style.id = STYLE_ID
    style.textContent = STYLES
    document.head.appendChild(style)
  }, [])
}
