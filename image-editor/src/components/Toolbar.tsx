import React from 'react';
import { 
  Pointer, 
  Move, 
  Crop, 
  Wand2, 
  Eraser, 
  Sparkle, 
  Paintbrush, 
  Layers, 
  Sliders, 
  RotateCw,
  Image as ImageIcon
} from 'lucide-react';
import { ToolType } from '../types/editor';

interface ToolbarProps {
  activeTool: ToolType;
  onSelectTool: (tool: ToolType) => void;
  hasRemovedBg: boolean;
}

interface ToolItem {
  id: ToolType;
  label: string;
  icon: React.ReactNode;
  shortcut: string;
  badge?: string;
}

export const Toolbar: React.FC<ToolbarProps> = ({
  activeTool,
  onSelectTool,
  hasRemovedBg,
}) => {
  const tools: ToolItem[] = [
    {
      id: 'select',
      label: 'Select / Inspect',
      icon: <Pointer size={19} />,
      shortcut: 'V',
    },
    {
      id: 'move',
      label: 'Pan Canvas',
      icon: <Move size={19} />,
      shortcut: 'H',
    },
    {
      id: 'crop',
      label: 'Crop & Resize',
      icon: <Crop size={19} />,
      shortcut: 'C',
    },
    {
      id: 'remove-bg',
      label: 'Remove Background',
      icon: <Wand2 size={19} />,
      shortcut: 'B',
      badge: 'AI',
    },
    {
      id: 'eraser',
      label: 'Eraser',
      icon: <Eraser size={19} />,
      shortcut: 'E',
    },
    {
      id: 'mark-remove',
      label: 'Mark & Remove',
      icon: <Sparkle size={19} />,
      shortcut: 'M',
      badge: 'NEW',
    },
    {
      id: 'restore',
      label: 'Restore Brush',
      icon: <Paintbrush size={19} />,
      shortcut: 'R',
    },
    {
      id: 'background',
      label: 'Background Options',
      icon: <Layers size={19} />,
      shortcut: 'G',
    },
    {
      id: 'adjustments',
      label: 'Filters & Adjustments',
      icon: <Sliders size={19} />,
      shortcut: 'A',
    },
    {
      id: 'transform',
      label: 'Rotate & Flip',
      icon: <RotateCw size={19} />,
      shortcut: 'T',
    },
  ];

  return (
    <aside className="editor-toolbar">
      <div className="toolbar-list">
        {tools.map((tool) => {
          const isActive = activeTool === tool.id;
          return (
            <button
              key={tool.id}
              className={`toolbar-btn ${isActive ? 'active' : ''}`}
              onClick={() => onSelectTool(tool.id)}
              title={`${tool.label} (${tool.shortcut})`}
            >
              <div className="tool-icon-box">
                {tool.icon}
                {tool.badge && <span className="tool-badge">{tool.badge}</span>}
              </div>
              <span className="tool-label">{tool.label}</span>
              <span className="tool-tooltip">
                {tool.label} <kbd>{tool.shortcut}</kbd>
              </span>
            </button>
          );
        })}
      </div>
    </aside>
  );
};
