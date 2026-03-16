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

  constructor(ws, opts) {
    this.#ws = ws;
    this.#opts = opts || {};
  }

  [Symbol.iterator]() {
    return this;
  }

  next() {
    while (this.#cursor >= this.#buffer.length) {
      if (this.#done) {
        return { done: true, value: undefined };
      }

      this.#buffer = this.#ws.getRowsBatch(this.#nextRow, BATCH_SIZE, this.#opts);
      this.#cursor = 0;

      if (this.#buffer.length === 0) {
        this.#done = true;
        return { done: true, value: undefined };
      }

      this.#nextRow = this.#buffer[this.#buffer.length - 1].index + 1;
    }

    return { done: false, value: this.#buffer[this.#cursor++] };
  }
}

const NativeWorksheet = native.Worksheet;

NativeWorksheet.prototype.iterateRows = function iterateRows(opts) {
  return new RowIterator(this, opts);
};

module.exports = native;
module.exports.RowIterator = RowIterator;
