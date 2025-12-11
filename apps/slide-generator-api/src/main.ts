
import { generateSlides, doPost, doGet } from './api';
import { GasSlidesApiTester } from './infrastructure/gas/GasSlidesApiTester';

// Globally expose the entry points for Google Apps Script
// This follows the standard pattern for GAS libraries built with modern tooling
// where we must explicitly attach functions to the global scope.

(global as any).generateSlides = generateSlides;
(global as any).doPost = doPost;
(global as any).doGet = doGet;

(global as any).testSlidesApi = () => {
    const pres = SlidesApp.create('API Prototype Test ' + new Date().toLocaleString());
    const id = pres.getId();
    Logger.log('Created Temp Presentation: ' + pres.getUrl());

    const tester = new GasSlidesApiTester();
    const result = tester.runPrototype(id);
    Logger.log(result);
};
