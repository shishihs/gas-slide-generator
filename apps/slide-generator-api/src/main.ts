
import { generateSlides, doPost, doGet } from './api';

// Globally expose the entry points for Google Apps Script
// This follows the standard pattern for GAS libraries built with modern tooling
// where we must explicitly attach functions to the global scope.

(global as any).generateSlides = generateSlides;
(global as any).doPost = doPost;
(global as any).doGet = doGet;
