import { parsePpt } from './parser.js';

self.onmessage = event => {
  try { self.postMessage({ presentation: parsePpt(event.data) }); }
  catch (error) { self.postMessage({ error: error.message }); }
};
