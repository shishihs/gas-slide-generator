module.exports = {
    preset: 'ts-jest',
    testEnvironment: 'node',
    testMatch: ['**/src/**/*.test.ts'],
    fontFamily: undefined, // Mock for canvas/chart related errors if any
};
