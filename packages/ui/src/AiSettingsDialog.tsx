import { useEffect, useId, useMemo, useRef, useState } from 'react'
import {
  AI_PROVIDERS,
  providerRequiresApiKey,
  type AiModelPreset,
  type AiProviderId,
  type AiProviderMeta,
  type AiProviderProbeRequest,
  type AiSettings,
  type CliStatus,
  type ModelListResult,
  type ProviderProbeResult,
} from '@genoffice/ai-provider'
import type { Lang } from '@genoffice/i18n'
import { tAiSettings } from './ai-settings-strings'

export interface AiSettingsDialogProps {
  /** current settings as the main process reports them (API keys already blanked) */
  settings: AiSettings
  lang: Lang
  /** persist and close; the caller sends this straight to `setAiSettings` */
  onSave: (settings: AiSettings) => void
  onClose: () => void
  /**
   * Main-process probes behind the Test and refresh buttons. Both are optional
   * so a host that has not wired the channels simply shows neither button.
   */
  onTest?: ((request: AiProviderProbeRequest) => Promise<ProviderProbeResult>) | undefined
  onListModels?: ((request: AiProviderProbeRequest) => Promise<ModelListResult>) | undefined
  /** ask the main process whether a local agent CLI backend is installed */
  onCliStatus?: ((provider: string) => Promise<CliStatus>) | undefined
  /**
   * Genspark sign-in state, so the hosted option can say whether it is usable.
   * Omitted where the host has no way to query it.
   */
  genspark?: { loggedIn: boolean; email?: string | undefined; onSignIn: () => void } | undefined
}

/**
 * Provider + model + API key editor, shared by all five apps.
 *
 * The dialog never receives a stored key (see `createAiSettingsStore`), only a
 * `hasApiKey` flag, so the key field starts empty with "a key is saved" next to
 * it: typing replaces the key, leaving it alone keeps it, and "remove key"
 * sends `hasApiKey: false` to delete it.
 *
 * Styling is injected by the component rather than left to the host apps: five
 * copies of the same CSS would drift, and the dialog is self-contained enough
 * that it does not need to match anything but itself.
 */
export function AiSettingsDialog({
  settings,
  lang,
  onSave,
  onClose,
  onTest,
  onListModels,
  onCliStatus,
  genspark,
}: AiSettingsDialogProps) {
  const t = (key: Parameters<typeof tAiSettings>[1], params?: Record<string, string>) =>
    tAiSettings(lang, key, params)
  useDialogStyles()

  const [draft, setDraft] = useState<AiSettings>(settings)
  const [error, setError] = useState<string | null>(null)
  const [revealKey, setRevealKey] = useState(false)
  /**
   * Outcome of the last Test, and models fetched from the provider. Both are
   * keyed by provider and kept only for the life of the dialog: a fetched list
   * is a snapshot of what the key could reach a moment ago, so re-opening the
   * dialog should ask again rather than show something stale.
   */
  const [probe, setProbe] = useState<Record<string, ProbeState | undefined>>({})
  const [fetched, setFetched] = useState<Record<string, string[] | undefined>>({})
  /** install state of the local agent CLIs, probed when their pane is shown */
  const [cli, setCli] = useState<Record<string, CliStatus | undefined>>({})
  const dialogRef = useRef<HTMLDivElement>(null)
  const modelInputId = useId()
  const modelListId = useId()

  const provider = draft.provider
  const meta = useMemo(
    () => AI_PROVIDERS.find((p) => p.id === provider) ?? AI_PROVIDERS[0]!,
    [provider],
  )
  const config = draft.providers[provider] ?? { apiKey: '', model: meta.defaultModel }
  const needsKey = providerRequiresApiKey(provider)
  const isCli = !!meta.isLocalCli
  const status = probe[provider]
  /** what the provider just told us it can serve, falling back to the built-in catalogue */
  const models = fetched[provider] ?? meta.models
  /**
   * Every backend can be asked for its catalogue except Genspark, whose proxy
   * publishes none. A local CLI answers over its own machine protocol rather
   * than HTTP, but from here that difference does not show.
   */
  const canListModels = !!onListModels && provider !== 'genspark'
  /** Test needs something to authenticate with: a typed key or a stored one */
  const canTest =
    !!onTest &&
    (!!config.model || !!meta.modelOptional) &&
    (!needsKey || !!config.apiKey || !!config.hasApiKey)

  // close on Escape, and start focus inside so the dialog is keyboard-usable
  useEffect(() => {
    dialogRef.current?.focus()
    const onKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose()
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [onClose])

  const patchProvider = (patch: Partial<AiSettings['providers'][AiProviderId]>) => {
    setError(null)
    // any edit invalidates the previous verdict — it was about the old values
    setProbe((prev) => ({ ...prev, [provider]: undefined }))
    setDraft((prev) => ({
      ...prev,
      providers: { ...prev.providers, [provider]: { ...prev.providers[provider]!, ...patch } },
    }))
  }

  const selectProvider = (id: AiProviderId) => {
    setError(null)
    setRevealKey(false)
    setDraft((prev) => ({ ...prev, provider: id }))
  }

  /** what the main process needs to reach this provider on the user's behalf */
  const probeRequest = (): AiProviderProbeRequest => ({
    provider,
    model: config.model,
    baseUrl: config.baseUrl,
    // sent only when the user typed a key that is not saved yet; otherwise the
    // main process uses the stored key and the stored endpoint
    apiKey: config.apiKey || undefined,
  })

  const runTest = () => {
    if (!onTest) return
    setProbe((prev) => ({ ...prev, [provider]: { status: 'busy' } }))
    void onTest(probeRequest())
      .then((result) =>
        setProbe((prev) => ({
          ...prev,
          [provider]: result.ok
            ? { status: 'ok' }
            : { status: 'failed', message: result.error ?? '' },
        })),
      )
      .catch((err: unknown) =>
        setProbe((prev) => ({
          ...prev,
          [provider]: {
            status: 'failed',
            message: err instanceof Error ? err.message : String(err),
          },
        })),
      )
  }

  /**
   * @param silent the automatic fetch on open. It still shows "Refreshing…" so
   *   the list filling in is explained, but a failure only falls back to the
   *   built-in catalogue: the user did not ask for this request, so an error
   *   banner they cannot act on would be noise. The refresh button reports.
   */
  const refreshModels = (silent = false) => {
    if (!onListModels) return
    setProbe((prev) => ({ ...prev, [provider]: { status: 'listing' } }))
    const fail = (message: string) =>
      setProbe((prev) => ({
        ...prev,
        [provider]: silent ? undefined : { status: 'failed', message },
      }))
    void onListModels(probeRequest())
      .then((result) => {
        if (!result.ok) {
          fail(result.error ?? '')
          return
        }
        const models = result.models ?? []
        if (!models.length) {
          fail(t('noModels'))
          return
        }
        setFetched((prev) => ({ ...prev, [provider]: models }))
        setProbe((prev) => ({ ...prev, [provider]: undefined }))
      })
      .catch((err: unknown) => fail(err instanceof Error ? err.message : String(err)))
  }

  /**
   * Fetch the catalogue as soon as a provider is shown, so the list is current
   * without the user having to think to refresh. Once per provider per dialog:
   * `attempted` also covers the failure case, which would otherwise retry on
   * every render.
   */
  // a CLI's presence is cheap to check and changes only if the user installs
  // something mid-session, so once per provider per dialog is enough
  useEffect(() => {
    if (!isCli || !onCliStatus || cli[provider]) return
    void onCliStatus(provider)
      .then((status) => setCli((prev) => ({ ...prev, [provider]: status })))
      .catch(() => setCli((prev) => ({ ...prev, [provider]: { installed: false } })))
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [provider, isCli])

  /**
   * Whether the fetch can succeed yet: an HTTP provider needs something to
   * authenticate with, a CLI needs to actually be on this machine (asking a
   * missing binary would only produce an error the user did not ask for).
   */
  const cliInstalled = cli[provider]?.installed === true
  const canAutoFetch = isCli ? cliInstalled : needsKey && (!!config.apiKey || !!config.hasApiKey)

  const attempted = useRef(new Set<AiProviderId>())
  useEffect(() => {
    if (!canListModels || !canAutoFetch) return
    if (fetched[provider] || attempted.current.has(provider)) return
    attempted.current.add(provider)
    refreshModels(true)
    // refreshModels closes over this render's provider/config; re-running on
    // anything else would just repeat a fetch `attempted` already guards
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [provider, config.apiKey, config.hasApiKey, cliInstalled])

  const save = () => {
    if (!config.model.trim() && !meta.modelOptional) return setError(t('errModel'))
    if (needsKey && !config.apiKey && !config.hasApiKey) return setError(t('errApiKey'))
    if (meta.needsBaseUrl && !config.baseUrl?.trim()) return setError(t('errBaseUrl'))
    onSave(draft)
  }

  return (
    <div className="gso-ai-settings-backdrop" onPointerDown={onClose}>
      <div
        className="gso-ai-settings"
        role="dialog"
        aria-modal="true"
        aria-label={t('title')}
        tabIndex={-1}
        ref={dialogRef}
        onPointerDown={(e) => e.stopPropagation()}
      >
        <div className="gso-ai-settings-head">
          <h2>{t('title')}</h2>
          <button className="gso-ai-settings-x" onClick={onClose} aria-label={t('cancel')}>
            ×
          </button>
        </div>

        <div className="gso-ai-settings-body">
          <nav className="gso-ai-settings-rail" aria-label={t('provider')}>
            {AI_PROVIDERS.map((option) => (
              <button
                key={option.id}
                className={`gso-ai-settings-rail-item${option.id === provider ? ' active' : ''}`}
                aria-current={option.id === provider}
                onClick={() => selectProvider(option.id)}
              >
                <span>{option.label}</span>
                {draft.providers[option.id]?.hasApiKey && (
                  <span className="gso-ai-settings-dot" aria-hidden="true" />
                )}
              </button>
            ))}
          </nav>

          <div className="gso-ai-settings-pane">
            {provider === 'genspark' ? (
              <GensparkPane genspark={genspark} t={t} />
            ) : isCli ? (
              <CliPane label={meta.label} status={cli[provider]} t={t} />
            ) : (
              <label className="gso-ai-settings-field">
                <span className="gso-ai-settings-label">{t('apiKey')}</span>
                <span className="gso-ai-settings-key-row">
                  <input
                    type={revealKey ? 'text' : 'password'}
                    value={config.apiKey}
                    autoComplete="off"
                    spellCheck={false}
                    placeholder={meta.keyPlaceholder}
                    onChange={(e) =>
                      patchProvider({ apiKey: e.target.value, hasApiKey: undefined })
                    }
                  />
                  <button
                    type="button"
                    className="gso-ai-settings-ghost"
                    aria-pressed={revealKey}
                    onClick={() => setRevealKey((v) => !v)}
                  >
                    {revealKey ? '🙈' : '👁'}
                  </button>
                </span>
                {config.hasApiKey && !config.apiKey && (
                  <span className="gso-ai-settings-hint gso-ai-settings-saved">
                    {t('apiKeyKeep')}
                    <button
                      type="button"
                      className="gso-ai-settings-link"
                      onClick={() => patchProvider({ apiKey: '', hasApiKey: false })}
                    >
                      {t('removeKey')}
                    </button>
                  </span>
                )}
              </label>
            )}

            <div className="gso-ai-settings-field">
              <label className="gso-ai-settings-label" htmlFor={modelInputId}>
                {t('model')}
              </label>
              <ModelCombo
                inputId={modelInputId}
                listId={modelListId}
                value={config.model}
                models={models}
                label={t('model')}
                onChange={(model) => patchProvider({ model })}
                onRefresh={canListModels ? () => refreshModels() : undefined}
                noMatchesLabel={t('noMatches')}
                refreshing={status?.status === 'listing'}
                refreshLabel={t('refresh')}
              />
              <span className="gso-ai-settings-hint">{t('modelHint')}</span>
            </div>

            <PresetEditor
              presets={draft.presets ?? []}
              provider={provider}
              model={config.model}
              t={t}
              onChange={(presets) => setDraft((prev) => ({ ...prev, presets }))}
              onApply={(model) => patchProvider({ model })}
            />

            {meta.needsBaseUrl && (
              <label className="gso-ai-settings-field">
                <span className="gso-ai-settings-label">{t('baseUrl')}</span>
                <input
                  type="text"
                  value={config.baseUrl ?? ''}
                  spellCheck={false}
                  placeholder="https://example.com/v1"
                  onChange={(e) => patchProvider({ baseUrl: e.target.value })}
                />
                <span className="gso-ai-settings-hint">{t('baseUrlHint')}</span>
              </label>
            )}

            {needsKey && <p className="gso-ai-settings-note">{t('storageNote')}</p>}
          </div>
        </div>

        <div className="gso-ai-settings-foot">
          {error ? (
            <span className="gso-ai-settings-error" role="alert">
              {error}
            </span>
          ) : (
            <ProbeStatus status={status} t={t} />
          )}
          {onTest && (
            <button
              className="gso-ai-settings-ghost"
              onClick={runTest}
              disabled={!canTest || status?.status === 'busy' || status?.status === 'listing'}
            >
              {status?.status === 'busy' ? t('testing') : t('test')}
            </button>
          )}
          <button className="gso-ai-settings-ghost" onClick={onClose}>
            {t('cancel')}
          </button>
          <button className="gso-ai-settings-primary" onClick={save}>
            {t('save')}
          </button>
        </div>
      </div>
    </div>
  )
}

/** ref callback that brings the selected option into view as the list mounts */
function scrollIntoView(el: HTMLElement | null): void {
  el?.scrollIntoView({ block: 'nearest' })
}

/** in-flight or finished outcome of a Test / model refresh, per provider */
interface ProbeState {
  status: 'busy' | 'listing' | 'ok' | 'failed'
  message?: string
}

/**
 * Result line in the footer. Provider errors are shown verbatim — they come
 * from the vendor and say far more than any phrase this dialog could invent
 * ("model X does not exist", "credit balance too low").
 */
function ProbeStatus({
  status,
  t,
}: {
  status: ProbeState | undefined
  t: (key: 'testOk' | 'refreshing') => string
}) {
  if (!status) return null
  if (status.status === 'ok') {
    return (
      <span className="gso-ai-settings-ok" role="status">
        ✓ {t('testOk')}
      </span>
    )
  }
  if (status.status === 'listing') {
    return (
      <span className="gso-ai-settings-muted" role="status">
        {t('refreshing')}
      </span>
    )
  }
  if (status.status === 'failed') {
    return (
      <span className="gso-ai-settings-error" role="alert" title={status.message}>
        {status.message}
      </span>
    )
  }
  return null
}

/**
 * Model picker: a typeahead over the provider's catalogue that still accepts
 * any id the provider will take (a new release, a fine-tune, a self-hosted
 * name).
 *
 * Deliberately not a `<datalist>`. The browser filters datalist suggestions
 * against whatever is already in the input, and this field is always prefilled
 * with the current model — so the list would collapse to the one entry the user
 * is already on and the rest would be unreachable. Here filtering starts only
 * once the user actually types: opening the list from the chevron always shows
 * everything, and a fetched catalogue runs to a hundred entries, so typing
 * "5.4" or "haiku" is the fast way through it.
 */
function ModelCombo({
  inputId,
  listId,
  value,
  models,
  label,
  onChange,
  onRefresh,
  refreshing,
  refreshLabel,
  noMatchesLabel,
}: {
  inputId: string
  listId: string
  value: string
  models: string[]
  label: string
  onChange: (model: string) => void
  /** re-read the catalogue from the provider; omitted where that is not possible */
  onRefresh?: (() => void) | undefined
  refreshing?: boolean | undefined
  refreshLabel: string
  noMatchesLabel: string
}) {
  // The list is positioned fixed against the input's measured rect: the
  // settings pane scrolls, and an absolutely-positioned popup would be clipped
  // at its edge (the shell's language flyout escapes its sidebar the same way).
  const [rect, setRect] = useState<{ left: number; top: number; width: number } | null>(null)
  /**
   * What the user has typed since opening the list, or null when they have not
   * typed at all. Kept separate from `value` so that a prefilled model does not
   * filter the list down to itself — the datalist trap this component exists to
   * avoid.
   */
  const [query, setQuery] = useState<string | null>(null)
  const open = rect !== null
  const wrapRef = useRef<HTMLDivElement>(null)
  const inputRef = useRef<HTMLInputElement>(null)

  const needle = query?.trim().toLowerCase() ?? ''
  const visible = needle ? models.filter((m) => m.toLowerCase().includes(needle)) : models

  const openList = (filtering = false) => {
    if (!filtering) setQuery(null)
    const box = inputRef.current?.getBoundingClientRect()
    if (box) setRect({ left: box.left, top: box.bottom + 4, width: box.width })
  }

  const closeList = () => {
    setRect(null)
    setQuery(null)
  }

  // close when the pointer goes down outside, and when anything scrolls: the
  // fixed popup would otherwise detach from the field it belongs to
  useEffect(() => {
    if (!open) return
    const onPointerDown = (e: PointerEvent) => {
      const target = e.target as Element | null
      if (wrapRef.current?.contains(target as Node)) return
      if (target?.closest?.('.gso-ai-settings-combo-list')) return
      closeList()
    }
    const onScroll = (e: Event) => {
      // the list scrolls its own options; only outside scrolls detach it
      const target = e.target as Element | null
      if (target instanceof Element && target.closest('.gso-ai-settings-combo-list')) return
      closeList()
    }
    document.addEventListener('pointerdown', onPointerDown, true)
    window.addEventListener('scroll', onScroll, true)
    window.addEventListener('resize', onScroll)
    return () => {
      document.removeEventListener('pointerdown', onPointerDown, true)
      window.removeEventListener('scroll', onScroll, true)
      window.removeEventListener('resize', onScroll)
    }
  }, [open])

  const choose = (model: string) => {
    onChange(model)
    closeList()
  }

  return (
    <div className="gso-ai-settings-combo" ref={wrapRef}>
      <div className="gso-ai-settings-key-row">
        <input
          id={inputId}
          ref={inputRef}
          type="text"
          role="combobox"
          aria-expanded={open}
          aria-controls={listId}
          aria-autocomplete="list"
          value={value}
          spellCheck={false}
          onChange={(e) => {
            onChange(e.target.value)
            // typing narrows the list, and opens it if it was closed
            setQuery(e.target.value)
            if (models.length) openList(true)
          }}
          onKeyDown={(e) => {
            if (e.key === 'ArrowDown' && models.length) {
              e.preventDefault()
              openList()
            } else if (e.key === 'Escape' && open) {
              // the dialog closes on Escape; while the list is open it should
              // only close the list
              e.stopPropagation()
              closeList()
            }
          }}
        />
        {models.length > 0 && (
          <button
            type="button"
            className="gso-ai-settings-ghost gso-ai-settings-combo-toggle"
            aria-label={label}
            aria-expanded={open}
            tabIndex={-1}
            onClick={() => (open ? closeList() : openList())}
          >
            <svg width="12" height="12" viewBox="0 0 12 12" aria-hidden="true">
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
        )}
        {onRefresh && (
          <button
            type="button"
            className="gso-ai-settings-ghost gso-ai-settings-combo-toggle"
            aria-label={refreshLabel}
            title={refreshLabel}
            disabled={refreshing}
            onClick={onRefresh}
          >
            <svg
              className={refreshing ? 'gso-ai-settings-spin' : undefined}
              width="13"
              height="13"
              viewBox="0 0 16 16"
              aria-hidden="true"
            >
              <path
                d="M13.4 7a5.5 5.5 0 1 0-.4 3.4"
                stroke="currentColor"
                strokeWidth="1.4"
                strokeLinecap="round"
                fill="none"
              />
              <path
                d="M13.6 2.9v4.2H9.4"
                stroke="currentColor"
                strokeWidth="1.4"
                strokeLinecap="round"
                strokeLinejoin="round"
                fill="none"
              />
            </svg>
          </button>
        )}
      </div>
      {rect && models.length > 0 && (
        <ul
          className="gso-ai-settings-combo-list"
          id={listId}
          role="listbox"
          style={{ left: rect.left, top: rect.top, width: rect.width }}
        >
          {visible.length === 0 ? (
            // say so rather than vanishing: the typed id is still accepted, it
            // just is not one the provider advertises
            <li className="gso-ai-settings-combo-empty">{noMatchesLabel}</li>
          ) : (
            visible.map((model) => (
              <li key={model}>
                <button
                  type="button"
                  role="option"
                  aria-selected={model === value}
                  className={`gso-ai-settings-combo-item${model === value ? ' active' : ''}`}
                  // a fetched catalogue runs to a hundred entries, sorted, so
                  // the current model is rarely near the top — open on it
                  ref={model === value ? scrollIntoView : undefined}
                  onClick={() => choose(model)}
                >
                  {model}
                </button>
              </li>
            ))
          )}
        </ul>
      )}
    </div>
  )
}

/**
 * Saved provider+model pairings for the provider on screen.
 *
 * A provider slot holds a single model, so switching between two models of the
 * same provider otherwise means retyping the id. Naming the pairing turns that
 * into one click here and one row in the composer's picker.
 *
 * Presets live in the draft like every other field: nothing is written until
 * the dialog is saved.
 */
function PresetEditor({
  presets,
  provider,
  model,
  t,
  onChange,
  onApply,
}: {
  presets: AiModelPreset[]
  provider: AiProviderId
  model: string
  t: (key: 'presets' | 'savePreset' | 'presetName' | 'deletePreset' | 'save' | 'cancel') => string
  onChange: (presets: AiModelPreset[]) => void
  /** load a preset's model into the field above */
  onApply: (model: string) => void
}) {
  const [name, setName] = useState<string | null>(null)
  const mine = presets.filter((preset) => preset.provider === provider)

  const commit = () => {
    const trimmed = (name ?? '').trim()
    if (!trimmed) return setName(null)
    // same name overwrites: the name is the preset's identity, and two rows
    // reading "fast" in the picker would be a puzzle rather than a choice
    onChange([
      ...presets.filter((preset) => preset.name !== trimmed),
      { name: trimmed, provider, model },
    ])
    setName(null)
  }

  return (
    <div className="gso-ai-settings-field">
      <span className="gso-ai-settings-label">{t('presets')}</span>
      <div className="gso-ai-settings-chips">
        {mine.map((preset) => (
          <span key={preset.name} className="gso-ai-settings-chip">
            <button
              type="button"
              className="gso-ai-settings-chip-main"
              title={preset.model}
              onClick={() => onApply(preset.model)}
            >
              {preset.name}
            </button>
            <button
              type="button"
              className="gso-ai-settings-chip-x"
              aria-label={`${t('deletePreset')} ${preset.name}`}
              title={t('deletePreset')}
              onClick={() => onChange(presets.filter((other) => other.name !== preset.name))}
            >
              ×
            </button>
          </span>
        ))}
        {name === null ? (
          <button
            type="button"
            className="gso-ai-settings-link gso-ai-settings-add-preset"
            onClick={() => setName('')}
          >
            + {t('savePreset')}
          </button>
        ) : (
          <span className="gso-ai-settings-key-row gso-ai-settings-preset-new">
            <input
              type="text"
              value={name}
              autoFocus
              maxLength={40}
              placeholder={t('presetName')}
              onChange={(e) => setName(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') {
                  e.preventDefault()
                  commit()
                } else if (e.key === 'Escape') {
                  // the dialog also closes on Escape; here it should only
                  // abandon the half-typed name
                  e.stopPropagation()
                  setName(null)
                }
              }}
            />
            <button type="button" className="gso-ai-settings-ghost" onClick={commit}>
              {t('save')}
            </button>
          </span>
        )}
      </div>
    </div>
  )
}

/**
 * Stands in for the API key field on a locally installed agent CLI. There is
 * no key to enter — what matters is whether the binary is on this machine, and
 * whether it is actually logged in, which only the Test button can answer.
 */
function CliPane({
  label,
  status,
  t,
}: {
  label: string
  status: CliStatus | undefined
  t: (key: 'cliInstalled' | 'cliMissing' | 'cliNote', params?: Record<string, string>) => string
}) {
  return (
    <div className="gso-ai-settings-field">
      <p className="gso-ai-settings-note">{t('cliNote', { name: label })}</p>
      {status &&
        (status.installed ? (
          <span className="gso-ai-settings-ok" role="status">
            ✓ {t('cliInstalled')}
            {status.version ? ` · ${status.version}` : ''}
            {status.path ? (
              <span className="gso-ai-settings-hint" title={status.path}>
                {status.path}
              </span>
            ) : null}
          </span>
        ) : (
          <span className="gso-ai-settings-error" role="alert">
            {t('cliMissing', { name: label })}
          </span>
        ))}
    </div>
  )
}

function GensparkPane({
  genspark,
  t,
}: {
  genspark: AiSettingsDialogProps['genspark']
  t: (
    key: 'gensparkHosted' | 'signIn' | 'signedInAs' | 'notSignedIn',
    params?: Record<string, string>,
  ) => string
}) {
  return (
    <div className="gso-ai-settings-field">
      <p className="gso-ai-settings-note">{t('gensparkHosted')}</p>
      {genspark && (
        <span className="gso-ai-settings-key-row">
          <span className="gso-ai-settings-status">
            {genspark.loggedIn
              ? t('signedInAs', { email: genspark.email ?? 'Genspark' })
              : t('notSignedIn')}
          </span>
          {!genspark.loggedIn && (
            <button type="button" className="gso-ai-settings-ghost" onClick={genspark.onSignIn}>
              {t('signIn')}
            </button>
          )}
        </span>
      )}
    </div>
  )
}

/** provider metadata re-exported so hosts can label a menu entry without a second import */
export type { AiProviderMeta }

// ── styles ───────────────────────────────────────────────
// Injected once per document. The apps have no shared stylesheet and no dark
// theme, so a light-theme block co-located with the markup is the whole story.

const STYLE_ID = 'gso-ai-settings-styles'

const STYLES = `
.gso-ai-settings-backdrop {
  position: fixed; inset: 0; z-index: 9000;
  display: flex; align-items: center; justify-content: center;
  background: rgba(17, 20, 24, 0.38);
}
.gso-ai-settings {
  width: min(620px, calc(100vw - 32px));
  max-height: calc(100vh - 64px);
  display: flex; flex-direction: column;
  background: #fff; color: #1b1f24;
  border-radius: 12px;
  box-shadow: 0 18px 48px rgba(15, 20, 30, 0.28);
  font: 13px/1.5 system-ui, -apple-system, "Segoe UI", sans-serif;
  outline: none;
}
.gso-ai-settings-head {
  display: flex; align-items: center; justify-content: space-between;
  padding: 14px 16px; border-bottom: 1px solid #e8eaee;
}
.gso-ai-settings-head h2 { margin: 0; font-size: 14px; font-weight: 600; }
.gso-ai-settings-x {
  border: 0; background: none; cursor: pointer;
  font-size: 20px; line-height: 1; color: #6b7280; padding: 0 4px;
}
.gso-ai-settings-x:hover { color: #1b1f24; }
.gso-ai-settings-body { display: flex; min-height: 0; flex: 1; }
.gso-ai-settings-rail {
  width: 148px; flex: none; padding: 10px 8px;
  border-right: 1px solid #e8eaee; overflow-y: auto;
}
.gso-ai-settings-rail-item {
  display: flex; align-items: center; gap: 6px; width: 100%;
  padding: 7px 10px; margin-bottom: 2px;
  border: 0; border-radius: 6px; background: none; cursor: pointer;
  font: inherit; text-align: start; color: #374151;
}
.gso-ai-settings-rail-item:hover { background: #f2f4f7; }
.gso-ai-settings-rail-item.active { background: #eaf1fe; color: #1a54c9; font-weight: 600; }
.gso-ai-settings-dot {
  width: 6px; height: 6px; border-radius: 50%;
  background: #16a34a; margin-inline-start: auto;
}
.gso-ai-settings-pane { flex: 1; min-width: 0; padding: 16px; overflow-y: auto; }
.gso-ai-settings-field { display: block; margin-bottom: 16px; }
.gso-ai-settings-label { display: block; margin-bottom: 5px; font-weight: 600; }
.gso-ai-settings-field input {
  width: 100%; box-sizing: border-box;
  padding: 7px 9px; border: 1px solid #d3d7de; border-radius: 6px;
  font: inherit; color: inherit; background: #fff;
}
.gso-ai-settings-field input:focus {
  outline: none; border-color: #3b82f6; box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.16);
}
.gso-ai-settings-key-row { display: flex; align-items: center; gap: 8px; }
/* fixed, not absolute: the settings pane scrolls and would clip the list */
.gso-ai-settings-combo-list {
  position: fixed; z-index: 9001;
  max-height: 210px; overflow-y: auto;
  margin: 0; padding: 4px; list-style: none;
  background: #fff; border: 1px solid #d3d7de; border-radius: 6px;
  box-shadow: 0 8px 20px rgba(15, 20, 30, 0.14);
}
.gso-ai-settings-combo-item {
  display: block; width: 100%; padding: 6px 9px;
  border: 0; border-radius: 4px; background: none; cursor: pointer;
  font: inherit; text-align: start; color: #374151;
}
.gso-ai-settings-combo-item:hover { background: #f2f4f7; }
.gso-ai-settings-combo-empty { padding: 6px 9px; color: #6b7280; }
.gso-ai-settings-combo-item.active { background: #eaf1fe; color: #1a54c9; font-weight: 600; }
.gso-ai-settings-chips { display: flex; flex-wrap: wrap; align-items: center; gap: 6px; }
.gso-ai-settings-chip {
  display: inline-flex; align-items: center;
  border: 1px solid #d3d7de; border-radius: 999px; background: #fff;
}
.gso-ai-settings-chip-main {
  border: 0; background: none; cursor: pointer; font: inherit;
  padding: 3px 4px 3px 10px; color: #374151;
}
.gso-ai-settings-chip-x {
  border: 0; background: none; cursor: pointer; font: inherit;
  padding: 3px 8px 3px 2px; color: #9aa1ac; line-height: 1;
}
.gso-ai-settings-chip:hover { background: #f2f4f7; }
.gso-ai-settings-chip-x:hover { color: #b91c1c; }
.gso-ai-settings-add-preset { margin-inline-start: 0; color: #1a54c9; text-decoration: none; }
.gso-ai-settings-add-preset:hover { text-decoration: underline; }
.gso-ai-settings-preset-new { flex: 1; min-width: 180px; }
.gso-ai-settings-status { flex: 1; min-width: 0; color: #4b5563; }
.gso-ai-settings-hint {
  display: block; margin-top: 5px; color: #6b7280; font-size: 12px;
}
.gso-ai-settings-saved { color: #15803d; }
.gso-ai-settings-link {
  border: 0; background: none; padding: 0; margin-inline-start: 8px;
  color: #b91c1c; cursor: pointer; font: inherit; text-decoration: underline;
}
.gso-ai-settings-note {
  margin: 0 0 12px; color: #6b7280; font-size: 12px;
}
.gso-ai-settings-foot {
  display: flex; align-items: center; justify-content: flex-end; gap: 8px;
  padding: 12px 16px; border-top: 1px solid #e8eaee;
}
.gso-ai-settings-error,
.gso-ai-settings-ok,
.gso-ai-settings-muted {
  margin-inline-end: auto; min-width: 0;
  overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
}
.gso-ai-settings-error { color: #b91c1c; }
.gso-ai-settings-ok { color: #15803d; font-weight: 600; }
.gso-ai-settings-field .gso-ai-settings-ok,
.gso-ai-settings-field .gso-ai-settings-error {
  display: block; margin-inline-end: 0; white-space: normal;
}
.gso-ai-settings-field .gso-ai-settings-ok .gso-ai-settings-hint {
  font-weight: 400; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
}
.gso-ai-settings-muted { color: #6b7280; }
.gso-ai-settings-ghost:disabled,
.gso-ai-settings-primary:disabled { opacity: 0.55; cursor: default; }
@keyframes gso-ai-settings-spin { to { transform: rotate(360deg); } }
.gso-ai-settings-spin { animation: gso-ai-settings-spin 0.9s linear infinite; transform-origin: 50% 50%; }
@media (prefers-reduced-motion: reduce) { .gso-ai-settings-spin { animation: none; } }
.gso-ai-settings-ghost,
.gso-ai-settings-primary {
  padding: 6px 14px; border-radius: 6px; font: inherit; cursor: pointer;
}
.gso-ai-settings-ghost { border: 1px solid #d3d7de; background: #fff; color: #374151; }
.gso-ai-settings-ghost:hover { background: #f2f4f7; }
.gso-ai-settings-primary { border: 1px solid #1a54c9; background: #1a54c9; color: #fff; }
.gso-ai-settings-primary:hover { background: #1747ad; }
/* must follow .gso-ai-settings-ghost: same specificity, and the ghost's
   horizontal padding would otherwise squash the chevron to a sliver */
.gso-ai-settings-combo-toggle {
  flex: none; display: flex; align-items: center; justify-content: center;
  width: 32px; height: 32px; padding: 0; color: #374151;
}
`

function useDialogStyles(): void {
  useEffect(() => {
    if (document.getElementById(STYLE_ID)) return
    const style = document.createElement('style')
    style.id = STYLE_ID
    style.textContent = STYLES
    document.head.appendChild(style)
  }, [])
}
