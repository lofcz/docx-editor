import { useEffect, useLayoutEffect, useRef, useState, type ReactNode } from 'react';
import type { BlockViewProps } from './react-block-layer';

function Chrome({ props, children }: { props: BlockViewProps; children: ReactNode }) {
  const ref = useRef<HTMLDivElement>(null);
  const resizeRef = useRef(props.resize);
  useLayoutEffect(() => {
    resizeRef.current = props.resize;
  });
  const hasHeader = Boolean(props.header);
  useEffect(() => {
    const element = ref.current;
    if (!element) return;
    const observer = new ResizeObserver(() =>
      resizeRef.current(element.getBoundingClientRect().height)
    );
    observer.observe(element);
    return () => observer.disconnect();
  }, [hasHeader]);
  const { header } = props;
  if (!header) return null;
  return (
    <div
      ref={ref}
      className="react-header"
      contentEditable={false}
      data-react-header={props.block.type}
      style={{ left: header.x, top: header.y, width: header.width }}
      onPointerDownCapture={(e) => e.stopPropagation()}
    >
      {children}
    </div>
  );
}
function Frame({ props, children }: { props: BlockViewProps; children: ReactNode }) {
  return (
    <section
      data-react-block={props.block.type}
      data-block-tag={props.block.tag}
      style={{ display: 'contents' }}
    >
      <div
        className={`react-frame ${props.block.type}`}
        contentEditable={false}
        style={{
          left: props.box.x - 5,
          top: props.box.y - 3,
          width: props.box.width + 10,
          height: props.box.height + 6,
        }}
      />
      {children}
      {props.children}
    </section>
  );
}
export function PhaseBlock(props: BlockViewProps) {
  return (
    <Frame props={props}>
      <Chrome props={props}>
        <strong>{props.block.data.name}</strong>
        <span className="badge">React container · native Word children</span>
      </Chrome>
    </Frame>
  );
}
export function ActivityBlock(props: BlockViewProps) {
  const [clicks, setClicks] = useState(0);
  const data = props.block.data;
  return (
    <Frame props={props}>
      <Chrome props={props}>
        <div className="activity-row">
          <strong>{data.name}</strong>
          <label>
            Minutes{' '}
            <input
              aria-label="Activity duration"
              type="number"
              value={data.duration ?? 0}
              onChange={(e) => props.update({ ...data, duration: Number(e.target.value) })}
            />
          </label>
        </div>
        <button onClick={() => setClicks((n) => n + 1)}>Local React state: {clicks}</button>
      </Chrome>
    </Frame>
  );
}
export function QuizBlock(props: BlockViewProps) {
  const data = props.block.data;
  return (
    <Frame props={props}>
      <Chrome props={props}>
        <strong>{data.name}</strong>
        <svg
          width="100%"
          height="42"
          viewBox="0 0 300 42"
          role="img"
          aria-label="Experiment diagram"
        >
          <path
            d="M10 35 L70 25 L130 30 L190 10 L260 5"
            fill="none"
            stroke="currentColor"
            strokeWidth="3"
          />
        </svg>
        <label>
          Prediction{' '}
          <select
            aria-label="Prediction"
            value={data.answer ?? ''}
            onChange={(e) => props.update({ ...data, answer: e.target.value })}
          >
            <option value="">Choose</option>
            <option value="warmer">Warmer</option>
            <option value="cooler">Cooler</option>
          </select>
        </label>
        <button onClick={() => props.update({ ...data, expanded: !data.expanded })}>
          {data.expanded ? 'Hide explanation' : 'Show explanation'}
        </button>
        {data.expanded ? (
          <div className="quiz-explanation">
            <p>
              This arbitrary React subtree changes its height. The measured height updates native
              Word line spacing.
            </p>
            <textarea aria-label="Scratch notes" placeholder="Independent browser input" />
          </div>
        ) : null}
      </Chrome>
    </Frame>
  );
}
