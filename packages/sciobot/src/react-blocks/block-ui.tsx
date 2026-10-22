import type { HTMLAttributes } from 'react';

/**
 * Widget controls are host UI inside the pages. Their focus, events, and text selections stay
 * outside native document selection. Put editable native slots beside this wrapper.
 */
export function BlockUi({
  children,
  style,
  ...props
}: Omit<HTMLAttributes<HTMLDivElement>, 'contentEditable'>) {
  return (
    <div
      {...props}
      contentEditable={false}
      data-sciobot-block-ui=""
      data-docx-host-ui=""
      style={{ ...style, pointerEvents: 'auto' }}
    >
      {children}
    </div>
  );
}
