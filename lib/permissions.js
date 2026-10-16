/**
 * DSH's Full Access preset is a standing host grant, so a write does not
 * need a second per-call approval prompt in that exact state.
 *
 * Requiring both native signals is intentional: `never` by itself means
 * reject approval requests, while `danger-full-access` by itself still uses
 * the host's normal approval channel.
 */
export function nativeFullAccess(ctx, agent) {
  const session = agent?.session
  if (!session) return false
  try {
    const approval = typeof ctx?.get === 'function' ? ctx.get('approval', false) : ctx?.approval
    const sandbox = typeof ctx?.get === 'function' ? ctx.get('sandboxPolicy', false) : ctx?.sandboxPolicy
    return approval?.effectivePolicy?.(session) === 'never'
      && sandbox?.resolve?.({ session })?.mode === 'danger-full-access'
  } catch {
    return false
  }
}
