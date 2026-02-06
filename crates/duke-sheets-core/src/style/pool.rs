//! Style pool for deduplication

use super::Style;
use ahash::AHashMap;

/// Style pool for deduplicating styles
///
/// Excel files typically have many cells sharing the same style.
/// The style pool ensures each unique style is stored only once,
/// and cells reference styles by index.
#[derive(Debug)]
pub struct StylePool {
    /// All unique styles (index 0 is default)
    styles: Vec<Style>,
    /// Fast lookup for deduplication
    index_map: AHashMap<StyleKey, u32>,
}

/// Key for style lookup (hash-based)
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
struct StyleKey(u64);

impl StyleKey {
    fn from_style(style: &Style) -> Self {
        use std::hash::{Hash, Hasher};
        let mut hasher = ahash::AHasher::default();
        style.hash(&mut hasher);
        StyleKey(hasher.finish())
    }
}

impl StylePool {
    /// Create a new style pool with default style at index 0
    pub fn new() -> Self {
        let mut pool = Self {
            styles: Vec::with_capacity(64),
            index_map: AHashMap::with_capacity(64),
        };

        // Index 0 is always the default style
        let default = Style::default();
        let key = StyleKey::from_style(&default);
        pool.styles.push(default);
        pool.index_map.insert(key, 0);

        pool
    }

    /// Get or create a style, returning its index
    ///
    /// If an identical style already exists, returns its index.
    /// Otherwise, adds the style and returns the new index.
    pub fn get_or_insert(&mut self, style: Style) -> u32 {
        let key = StyleKey::from_style(&style);

        if let Some(&idx) = self.index_map.get(&key) {
            // Verify it's actually the same (hash collision check)
            if self.styles[idx as usize] == style {
                return idx;
            }
        }

        // Not found or collision, add new
        let idx = self.styles.len() as u32;
        self.index_map.insert(key, idx);
        self.styles.push(style);
        idx
    }

    /// Get a style by index
    pub fn get(&self, index: u32) -> Option<&Style> {
        self.styles.get(index as usize)
    }

    /// Get the default style (index 0)
    pub fn default_style(&self) -> &Style {
        &self.styles[0]
    }

    /// Get the number of styles
    pub fn len(&self) -> usize {
        self.styles.len()
    }

    /// Check if the pool is empty (only has default)
    pub fn is_empty(&self) -> bool {
        self.styles.len() <= 1
    }

    /// Iterate over all styles with their indices
    pub fn iter(&self) -> impl Iterator<Item = (u32, &Style)> {
        self.styles.iter().enumerate().map(|(i, s)| (i as u32, s))
    }

    /// Clear all styles except default
    pub fn clear(&mut self) {
        let default = self.styles[0].clone();
        self.styles.clear();
        self.index_map.clear();

        let key = StyleKey::from_style(&default);
        self.styles.push(default);
        self.index_map.insert(key, 0);
    }
}

impl Default for StylePool {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::style::{Color, FillStyle};

    #[test]
    fn test_default_style() {
        let pool = StylePool::new();
        assert_eq!(pool.len(), 1);
        assert_eq!(pool.get(0), Some(&Style::default()));
    }

    #[test]
    fn test_deduplication() {
        let mut pool = StylePool::new();

        let style1 = Style::new().bold(true);
        let style2 = Style::new().bold(true); // Same as style1
        let style3 = Style::new().italic(true); // Different

        let idx1 = pool.get_or_insert(style1);
        let idx2 = pool.get_or_insert(style2);
        let idx3 = pool.get_or_insert(style3);

        assert_eq!(idx1, idx2); // Same style, same index
        assert_ne!(idx1, idx3); // Different style, different index
        assert_eq!(pool.len(), 3); // default + 2 custom
    }

    #[test]
    fn test_complex_styles() {
        let mut pool = StylePool::new();

        let style = Style::new()
            .bold(true)
            .italic(true)
            .font_size(14.0)
            .fill_color(Color::RED);

        let idx = pool.get_or_insert(style.clone());
        assert!(idx > 0);
        assert_eq!(pool.get(idx), Some(&style));
    }
}
