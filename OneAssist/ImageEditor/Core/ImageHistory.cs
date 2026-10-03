using System;
using System.Collections.Generic;
using System.Drawing;

namespace OneAssist.ImageEditor.Core
{
    public class ImageHistory : IDisposable
    {
        private readonly int _maxSteps;
        private readonly List<HistoryStep> _undoStack = new List<HistoryStep>();
        private readonly List<HistoryStep> _redoStack = new List<HistoryStep>();

        public event EventHandler HistoryChanged;

        public bool CanUndo => _undoStack.Count > 0;
        public bool CanRedo => _redoStack.Count > 0;

        public ImageHistory(int maxSteps = 20)
        {
            _maxSteps = maxSteps;
        }

        public void Push(string description, Bitmap workingImage, ImageMask mask)
        {
            var step = new HistoryStep(description, workingImage, mask);
            _undoStack.Add(step);

            // Clear redo stack on new operation
            ClearRedoStack();

            // Prune excess undo steps
            while (_undoStack.Count > _maxSteps)
            {
                _undoStack[0].Dispose();
                _undoStack.RemoveAt(0);
            }

            HistoryChanged?.Invoke(this, EventArgs.Empty);
        }

        public HistoryStep PopUndo(Bitmap currentWorkingImage, ImageMask currentMask, string redoDescription)
        {
            if (!CanUndo) return null;

            int lastIdx = _undoStack.Count - 1;
            var stepToRestore = _undoStack[lastIdx];
            _undoStack.RemoveAt(lastIdx);

            // Save current state to redo stack
            _redoStack.Add(new HistoryStep(redoDescription ?? stepToRestore.Description, currentWorkingImage, currentMask));

            HistoryChanged?.Invoke(this, EventArgs.Empty);
            return stepToRestore;
        }

        public HistoryStep PopRedo(Bitmap currentWorkingImage, ImageMask currentMask, string undoDescription)
        {
            if (!CanRedo) return null;

            int lastIdx = _redoStack.Count - 1;
            var stepToRestore = _redoStack[lastIdx];
            _redoStack.RemoveAt(lastIdx);

            // Push current state back to undo stack
            _undoStack.Add(new HistoryStep(undoDescription ?? stepToRestore.Description, currentWorkingImage, currentMask));

            HistoryChanged?.Invoke(this, EventArgs.Empty);
            return stepToRestore;
        }

        private void ClearRedoStack()
        {
            foreach (var step in _redoStack)
            {
                step.Dispose();
            }
            _redoStack.Clear();
        }

        public void Clear()
        {
            foreach (var step in _undoStack) step.Dispose();
            _undoStack.Clear();

            ClearRedoStack();
            HistoryChanged?.Invoke(this, EventArgs.Empty);
        }

        public void Dispose()
        {
            Clear();
        }
    }
}
