//! Dependency tracking for formula calculation

use duke_sheets_core::CellAddress;
use std::collections::{HashMap, HashSet};

/// Unique key for a cell (sheet index + address)
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct CellKey {
    pub sheet: usize,
    pub row: u32,
    pub col: u16,
}

impl CellKey {
    /// Create a new cell key
    pub fn new(sheet: usize, row: u32, col: u16) -> Self {
        Self { sheet, row, col }
    }

    /// Create from sheet index and cell address
    pub fn from_address(sheet: usize, addr: &CellAddress) -> Self {
        Self::new(sheet, addr.row, addr.col)
    }
}

/// Dependency graph for formula cells
///
/// Tracks which cells depend on which other cells,
/// enabling efficient recalculation.
#[derive(Debug, Default)]
pub struct DependencyGraph {
    /// Cell → Cells that depend on it (dependents)
    dependents: HashMap<CellKey, HashSet<CellKey>>,
    /// Cell → Cells it depends on (precedents)
    precedents: HashMap<CellKey, HashSet<CellKey>>,
}

impl DependencyGraph {
    /// Create a new empty dependency graph
    pub fn new() -> Self {
        Self::default()
    }

    /// Add a dependency: dependent depends on precedent
    pub fn add_dependency(&mut self, precedent: CellKey, dependent: CellKey) {
        self.dependents
            .entry(precedent)
            .or_default()
            .insert(dependent);
        self.precedents
            .entry(dependent)
            .or_default()
            .insert(precedent);
    }

    /// Remove all dependencies for a cell
    pub fn clear_dependencies(&mut self, cell: CellKey) {
        // Remove from all precedents' dependents list
        if let Some(precedents) = self.precedents.remove(&cell) {
            for precedent in precedents {
                if let Some(deps) = self.dependents.get_mut(&precedent) {
                    deps.remove(&cell);
                }
            }
        }

        // Remove as a precedent for others
        if let Some(dependents) = self.dependents.remove(&cell) {
            for dependent in dependents {
                if let Some(precs) = self.precedents.get_mut(&dependent) {
                    precs.remove(&cell);
                }
            }
        }
    }

    /// Get cells that depend on the given cell
    pub fn get_dependents(&self, cell: CellKey) -> impl Iterator<Item = CellKey> + '_ {
        self.dependents
            .get(&cell)
            .into_iter()
            .flat_map(|set| set.iter().copied())
    }

    /// Get cells that the given cell depends on
    pub fn get_precedents(&self, cell: CellKey) -> impl Iterator<Item = CellKey> + '_ {
        self.precedents
            .get(&cell)
            .into_iter()
            .flat_map(|set| set.iter().copied())
    }

    /// Get all cells that need to be recalculated when the given cells change
    pub fn get_recalc_order(&self, changed: &[CellKey]) -> Vec<CellKey> {
        let mut result = Vec::new();
        let mut visited = HashSet::new();
        let mut in_stack = HashSet::new();

        for &cell in changed {
            self.topological_sort(cell, &mut result, &mut visited, &mut in_stack);
        }

        result
    }

    /// Topological sort helper (DFS)
    fn topological_sort(
        &self,
        cell: CellKey,
        result: &mut Vec<CellKey>,
        visited: &mut HashSet<CellKey>,
        in_stack: &mut HashSet<CellKey>,
    ) {
        if visited.contains(&cell) {
            return;
        }

        if in_stack.contains(&cell) {
            // Circular reference - skip (will be handled elsewhere)
            return;
        }

        in_stack.insert(cell);

        // Visit all dependents first
        if let Some(dependents) = self.dependents.get(&cell) {
            for &dependent in dependents {
                self.topological_sort(dependent, result, visited, in_stack);
            }
        }

        in_stack.remove(&cell);
        visited.insert(cell);
        result.push(cell);
    }

    /// Detect circular references involving a cell
    pub fn has_circular_reference(&self, cell: CellKey) -> bool {
        let mut visited = HashSet::new();
        let mut in_stack = HashSet::new();
        self.detect_cycle(cell, &mut visited, &mut in_stack)
    }

    fn detect_cycle(
        &self,
        cell: CellKey,
        visited: &mut HashSet<CellKey>,
        in_stack: &mut HashSet<CellKey>,
    ) -> bool {
        if in_stack.contains(&cell) {
            return true;
        }
        if visited.contains(&cell) {
            return false;
        }

        visited.insert(cell);
        in_stack.insert(cell);

        if let Some(precedents) = self.precedents.get(&cell) {
            for &precedent in precedents {
                if self.detect_cycle(precedent, visited, in_stack) {
                    return true;
                }
            }
        }

        in_stack.remove(&cell);
        false
    }

    /// Clear the entire graph
    pub fn clear(&mut self) {
        self.dependents.clear();
        self.precedents.clear();
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_add_dependency() {
        let mut graph = DependencyGraph::new();

        let a1 = CellKey::new(0, 0, 0);
        let b1 = CellKey::new(0, 0, 1);

        graph.add_dependency(a1, b1);

        assert!(graph.get_dependents(a1).any(|c| c == b1));
        assert!(graph.get_precedents(b1).any(|c| c == a1));
    }

    #[test]
    fn test_circular_reference() {
        let mut graph = DependencyGraph::new();

        let a1 = CellKey::new(0, 0, 0);
        let b1 = CellKey::new(0, 0, 1);
        let c1 = CellKey::new(0, 0, 2);

        // A1 -> B1 -> C1 -> A1 (circular)
        graph.add_dependency(a1, b1);
        graph.add_dependency(b1, c1);
        graph.add_dependency(c1, a1);

        assert!(graph.has_circular_reference(a1));
        assert!(graph.has_circular_reference(b1));
        assert!(graph.has_circular_reference(c1));
    }
}
