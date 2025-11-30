using System;
using System.Collections.Generic;

namespace WindowInspector.Models
{
    public class WindowConfig
    {
        public string WindowClass { get; set; } = string.Empty;
        public string WindowTitle { get; set; } = string.Empty;
        public List<InputPosition> InputPositions { get; set; } = new();
        public List<SavedTextItem> SavedTexts { get; set; } = new();
        public Dictionary<string, int> ImeMethods { get; set; } = new();
        public string? TargetProgramPath { get; set; }
        public bool AutoLaunch { get; set; }
        public bool IsExcelMode { get; set; }
        public List<string> ExcelCells { get; set; } = new();
        public List<CellGroup> ExcelCellGroups { get; set; } = new();
        public int ActiveCellGroupIndex { get; set; }
        
        // 地鼠分组加载设置
        public bool AutoLoadMoleGroups { get; set; } = false;
        public List<string> SelectedMoleGroups { get; set; } = new();
    }

    public class InputPosition
    {
        public int X { get; set; }
        public int Y { get; set; }
    }

    public class SavedTextItem
    {
        public string Name { get; set; } = string.Empty;
        public List<string> Texts { get; set; } = new();
        public List<string> Cases { get; set; } = new();
        public DateTime? LastFilledTime { get; set; }
        public bool FromExcel { get; set; }
    }

    public class CellGroup
    {
        public string Name { get; set; } = string.Empty;
        public List<string> Cells { get; set; } = new();
    }

    public class WindowPosition
    {
        public int X { get; set; }
        public int Y { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
    }
}
