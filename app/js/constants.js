window.APP_CONSTANTS = window.APP_CONSTANTS || {
  projectSchemaVersion: 1,
  modules: ['home', 'storm', 'condensate', 'vent', 'gas', 'wsfu', 'fixtureUnit', 'solar', 'duct', 'ductStatic', 'refrigerant'],
  defaults: {
    stormRainfallRate: 1,
    condensateQty: 1,
    wsfuDesignLengthFt: 100
  },
  messages: {
    ductPlaceholder: 'Enter airflow and sizing criteria to see live duct size recommendations.',
    ductStaticPlaceholder: 'Enter duct segment data to calculate total static pressure drop.',
    calculateFirst: 'Preview updates as you edit inputs. Click CALCULATE to commit to Results.'
  }
};
