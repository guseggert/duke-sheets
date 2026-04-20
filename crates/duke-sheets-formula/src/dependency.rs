//! Dependency tracking for formula calculation

use duke_sheets_core::CellAddress;
use std::cell::RefCell;
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
    circular_cells_cache: RefCell<Option<HashSet<CellKey>>>,
}

impl DependencyGraph {
    /// Create a new empty dependency graph
    pub fn new() -> Self {
        Self::default()
    }

    /// Add a dependency: dependent depends on precedent
    pub fn add_dependency(&mut self, precedent: CellKey, dependent: CellKey) {
        self.invalidate_circular_cache();
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
        self.invalidate_circular_cache();

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

    /// Detect circular references involving a cell (legacy per-cell API)
    pub fn has_circular_reference(&self, cell: CellKey) -> bool {
        self.ensure_circular_cells_cache();
        self.circular_cells_cache
            .borrow()
            .as_ref()
            .is_some_and(|circular_cells| circular_cells.contains(&cell))
    }

    /// Find ALL cells involved in circular references using Tarjan's SCC algorithm.
    ///
    /// Returns the set of cells that belong to any strongly connected component
    /// with 2+ members, plus self-referential cells with a direct self-edge.
    /// Runs in O(V+E) - single pass over the entire graph, much faster than
    /// per-cell cycle detection.
    pub fn find_circular_cells(&self) -> HashSet<CellKey> {
        self.ensure_circular_cells_cache();
        self.circular_cells_cache
            .borrow()
            .as_ref()
            .cloned()
            .unwrap_or_default()
    }

    fn invalidate_circular_cache(&self) {
        self.circular_cells_cache.replace(None);
    }

    fn ensure_circular_cells_cache(&self) {
        if self.circular_cells_cache.borrow().is_none() {
            let circular_cells = self.compute_circular_cells();
            self.circular_cells_cache.replace(Some(circular_cells));
        }
    }

    fn compute_circular_cells(&self) -> HashSet<CellKey> {
        // Collect all nodes that appear in the graph (as precedents or dependents)
        let mut all_nodes: HashSet<CellKey> = HashSet::new();
        for (&cell, deps) in &self.precedents {
            all_nodes.insert(cell);
            for &dep in deps {
                all_nodes.insert(dep);
            }
        }
        for (&cell, deps) in &self.dependents {
            all_nodes.insert(cell);
            for &dep in deps {
                all_nodes.insert(dep);
            }
        }

        let mut state = TarjanState {
            index_counter: 0,
            stack: Vec::new(),
            on_stack: HashSet::new(),
            indices: HashMap::new(),
            lowlinks: HashMap::new(),
            circular: HashSet::new(),
            precedents: &self.precedents,
        };

        for &node in &all_nodes {
            if !state.indices.contains_key(&node) {
                state.strongconnect(node);
            }
        }

        state.circular
    }

    /// Clear the entire graph
    pub fn clear(&mut self) {
        self.invalidate_circular_cache();
        self.dependents.clear();
        self.precedents.clear();
    }
}

/// Internal state for Tarjan's SCC algorithm (iterative to avoid stack overflow)
struct TarjanState<'a> {
    index_counter: u32,
    stack: Vec<CellKey>,
    on_stack: HashSet<CellKey>,
    indices: HashMap<CellKey, u32>,
    lowlinks: HashMap<CellKey, u32>,
    circular: HashSet<CellKey>,
    /// We traverse the precedent edges (cell → what it depends on)
    precedents: &'a HashMap<CellKey, HashSet<CellKey>>,
}

impl TarjanState<'_> {
    /// Iterative Tarjan's strongconnect to avoid stack overflow on deep graphs.
    fn strongconnect(&mut self, root: CellKey) {
        // Work stack: (node, iterator_state)
        // We use an explicit stack with frames to simulate recursion.
        // Each frame holds the node, an index into its neighbors, and whether
        // it has been initialized.

        // Collect neighbors into vecs for indexed iteration
        let mut neighbor_cache: HashMap<CellKey, Vec<CellKey>> = HashMap::new();

        struct Frame {
            node: CellKey,
            neighbor_idx: usize,
        }

        let mut work: Vec<Frame> = Vec::new();

        // Initialize root
        self.indices.insert(root, self.index_counter);
        self.lowlinks.insert(root, self.index_counter);
        self.index_counter += 1;
        self.stack.push(root);
        self.on_stack.insert(root);

        work.push(Frame {
            node: root,
            neighbor_idx: 0,
        });

        while let Some(frame) = work.last_mut() {
            let v = frame.node;

            // Get or build neighbor list for this node
            let neighbors = neighbor_cache
                .entry(v)
                .or_insert_with(|| {
                    self.precedents
                        .get(&v)
                        .map(|s| s.iter().copied().collect())
                        .unwrap_or_default()
                })
                .clone();

            if frame.neighbor_idx < neighbors.len() {
                let w = neighbors[frame.neighbor_idx];
                frame.neighbor_idx += 1;

                if !self.indices.contains_key(&w) {
                    // w not yet visited - "recurse"
                    self.indices.insert(w, self.index_counter);
                    self.lowlinks.insert(w, self.index_counter);
                    self.index_counter += 1;
                    self.stack.push(w);
                    self.on_stack.insert(w);

                    work.push(Frame {
                        node: w,
                        neighbor_idx: 0,
                    });
                } else if self.on_stack.contains(&w) {
                    // w is on stack - back edge
                    let w_index = self.indices[&w];
                    let v_lowlink = self.lowlinks.get_mut(&v).unwrap();
                    if w_index < *v_lowlink {
                        *v_lowlink = w_index;
                    }
                }
            } else {
                // Done processing all neighbors of v
                let v_lowlink = self.lowlinks[&v];
                let v_index = self.indices[&v];

                // Pop this frame
                work.pop();

                // Propagate lowlink to parent
                if let Some(parent_frame) = work.last() {
                    let parent = parent_frame.node;
                    let parent_lowlink = self.lowlinks.get_mut(&parent).unwrap();
                    if v_lowlink < *parent_lowlink {
                        *parent_lowlink = v_lowlink;
                    }
                }

                // If v is a root of an SCC, pop the SCC off the stack
                if v_lowlink == v_index {
                    let mut scc = Vec::new();
                    loop {
                        let w = self.stack.pop().unwrap();
                        self.on_stack.remove(&w);
                        scc.push(w);
                        if w == v {
                            break;
                        }
                    }
                    let has_self_loop = scc.len() == 1
                        && self
                            .precedents
                            .get(&v)
                            .is_some_and(|precedents| precedents.contains(&v));
                    if scc.len() > 1 || has_self_loop {
                        for cell in scc {
                            self.circular.insert(cell);
                        }
                    }
                }
            }
        }
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
        assert_eq!(graph.find_circular_cells(), HashSet::from([a1, b1, c1]));
    }

    #[test]
    fn test_find_circular_cells_basic() {
        let mut graph = DependencyGraph::new();

        let a1 = CellKey::new(0, 0, 0);
        let b1 = CellKey::new(0, 0, 1);
        let c1 = CellKey::new(0, 0, 2);
        let d1 = CellKey::new(0, 0, 3); // not in cycle

        // A1 -> B1 -> C1 -> A1 (circular)
        graph.add_dependency(a1, b1);
        graph.add_dependency(b1, c1);
        graph.add_dependency(c1, a1);
        // D1 -> A1 (not circular itself)
        graph.add_dependency(a1, d1);

        let circular = graph.find_circular_cells();
        assert!(circular.contains(&a1));
        assert!(circular.contains(&b1));
        assert!(circular.contains(&c1));
        assert!(!circular.contains(&d1));
    }

    #[test]
    fn test_find_circular_cells_no_cycles() {
        let mut graph = DependencyGraph::new();

        let a1 = CellKey::new(0, 0, 0);
        let b1 = CellKey::new(0, 0, 1);
        let c1 = CellKey::new(0, 0, 2);

        // Linear chain: A1 -> B1 -> C1
        graph.add_dependency(a1, b1);
        graph.add_dependency(b1, c1);

        let circular = graph.find_circular_cells();
        assert!(circular.is_empty());
    }

    #[test]
    fn test_find_circular_cells_self_reference() {
        let mut graph = DependencyGraph::new();

        let a1 = CellKey::new(0, 0, 0);
        // A1 depends on itself
        graph.add_dependency(a1, a1);

        assert!(graph.has_circular_reference(a1));
        assert_eq!(graph.find_circular_cells(), HashSet::from([a1]));
    }

    #[test]
    fn test_find_circular_cells_multiple_sccs() {
        let mut graph = DependencyGraph::new();

        let a1 = CellKey::new(0, 0, 0);
        let b1 = CellKey::new(0, 0, 1);
        let c1 = CellKey::new(0, 0, 2);
        let d1 = CellKey::new(0, 0, 3);

        // Cycle 1: A1 -> B1 -> A1
        graph.add_dependency(a1, b1);
        graph.add_dependency(b1, a1);

        // Cycle 2: C1 -> D1 -> C1
        graph.add_dependency(c1, d1);
        graph.add_dependency(d1, c1);

        let circular = graph.find_circular_cells();
        assert!(circular.contains(&a1));
        assert!(circular.contains(&b1));
        assert!(circular.contains(&c1));
        assert!(circular.contains(&d1));
    }
}
