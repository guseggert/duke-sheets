// @ts-nocheck
'use strict';

const native = require('./index.js');

const BATCH_SIZE = 1000;

class RowIterator {
  #ws;
  #opts;
  #buffer = [];
  #cursor = 0;
  #nextRow = 0;
  #done = false;
  #maxRow;

  constructor(ws, opts, maxRow) {
    this.#ws = ws;
    this.#opts = opts || {};
    this.#maxRow = maxRow;
  }

  [Symbol.iterator]() {
    return this;
  }

  next() {
    while (this.#cursor >= this.#buffer.length) {
      if (this.#done || this.#nextRow > this.#maxRow) {
        this.#done = true;
        return { done: true, value: undefined };
      }

      const batchSize = Math.min(BATCH_SIZE, this.#maxRow - this.#nextRow + 1);
      this.#buffer = this.#ws.getRowsBatch(this.#nextRow, batchSize, this.#opts);
      this.#cursor = 0;

      if (this.#buffer.length === 0) {
        this.#nextRow += batchSize;
        continue;
      }

      this.#nextRow = this.#buffer[this.#buffer.length - 1].index + 1;
    }

    return { done: false, value: this.#buffer[this.#cursor++] };
  }
}

const NativeWorksheet = native.Worksheet;

NativeWorksheet.prototype.iterateRows = function iterateRows(opts) {
  const range = this.usedRange;
  if (!range) return new RowIterator(this, opts, 0);
  return new RowIterator(this, opts, range.maxRow);
};

module.exports = native;
module.exports.RowIterator = RowIterator;
