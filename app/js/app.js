var APP_CONFIG = window.APP_CONSTANTS || {};
var APP_DEFAULTS = APP_CONFIG.defaults || {};
var APP_MESSAGES = APP_CONFIG.messages || {};
var APP_SCHEMA_VERSION = (typeof APP_CONFIG.projectSchemaVersion === 'number') ? APP_CONFIG.projectSchemaVersion : 1;
var APP_VERSION = '';
var pipeData = {
      "IPC": {
        label: "International Plumbing Code",
        vertical: [
          { size: 2, gpm: 34 }, { size: 3, gpm: 87 }, { size: 4, gpm: 180 }, { size: 5, gpm: 311 },
          { size: 6, gpm: 538 }, { size: 8, gpm: 1117 }, { size: 10, gpm: 2050 }, { size: 12, gpm: 3272 }, { size: 15, gpm: 5543 }
        ],
        horizontal: {
          "1/16 in/ft": [{ size: 2, gpm: 15 }, { size: 3, gpm: 39 }, { size: 4, gpm: 81 }, { size: 5, gpm: 117 }, { size: 6, gpm: 243 }, { size: 8, gpm: 505 }, { size: 10, gpm: 927 }, { size: 12, gpm: 1480 }, { size: 15, gpm: 2508 }],
          "1/8 in/ft":  [{ size: 2, gpm: 22 }, { size: 3, gpm: 55 }, { size: 4, gpm: 115 }, { size: 5, gpm: 165 }, { size: 6, gpm: 344 }, { size: 8, gpm: 714 }, { size: 10, gpm: 1311 }, { size: 12, gpm: 2093 }, { size: 15, gpm: 3546 }],
          "1/4 in/ft":  [{ size: 2, gpm: 31 }, { size: 3, gpm: 79 }, { size: 4, gpm: 163 }, { size: 5, gpm: 234 }, { size: 6, gpm: 487 }, { size: 8, gpm: 1010 }, { size: 10, gpm: 1855 }, { size: 12, gpm: 2960 }, { size: 15, gpm: 5016 }],
          "1/2 in/ft":  [{ size: 2, gpm: 44 }, { size: 3, gpm: 111 }, { size: 4, gpm: 231 }, { size: 5, gpm: 331 }, { size: 6, gpm: 689 }, { size: 8, gpm: 1429 }, { size: 10, gpm: 2623 }, { size: 12, gpm: 4187 }, { size: 15, gpm: 7093 }]
        }
      },
      "CPC": {
        label: "California Plumbing Code",
        vertical: [
          { size: 2, gpm: 30 }, { size: 3, gpm: 92 }, { size: 4, gpm: 192 }, { size: 5, gpm: 360 },
          { size: 6, gpm: 563 }, { size: 8, gpm: 1208 }
        ],
        horizontal: {
          "1/8 in/ft": [{ size: 3, gpm: 34 }, { size: 4, gpm: 78 }, { size: 5, gpm: 139 }, { size: 6, gpm: 222 }, { size: 8, gpm: 478 }, { size: 10, gpm: 860 }, { size: 12, gpm: 1384 }, { size: 15, gpm: 2473 }],
          "1/4 in/ft": [{ size: 3, gpm: 48 }, { size: 4, gpm: 110 }, { size: 5, gpm: 196 }, { size: 6, gpm: 314 }, { size: 8, gpm: 677 }, { size: 10, gpm: 1214 }, { size: 12, gpm: 1953 }, { size: 15, gpm: 3491 }],
          "1/2 in/ft": [{ size: 3, gpm: 68 }, { size: 4, gpm: 156 }, { size: 5, gpm: 278 }, { size: 6, gpm: 445 }, { size: 8, gpm: 956 }, { size: 10, gpm: 1721 }, { size: 12, gpm: 2768 }, { size: 15, gpm: 4946 }]
        }
      }
    };
    var latestPerAreaExport = null;
    var latestGasSnapshot = null;
    var latestCondensateSnapshot = null;
    var latestVentSnapshot = null;
    var latestWsfuSnapshot = null;
    var latestSolarSnapshot = null;
    var latestDuctSnapshot = null;
    var latestDuctStaticSnapshot = null;
    var latestRefrigerantSnapshot = null;
    var latestFixtureUnitSnapshot = null;
    var latestVentilationSnapshot = null;
    var fixtureUnitZones = [];
    var fixtureUnitZoneIdSeq = 0;
    var fixtureUnitRowIdSeq = 0;
    var refrigerantMultiRows = [];
    var refrigerantMultiRowIdSeq = 0;
    var ductLiveTimer = null;
    var fixtureUnitCalcTimer = null;
    var solarZipLookupTimer = null;
    var solarClimateManualOverride = false;
    var solarAutoCalcTimer = null;
    var activeModule = 'home';
    var activeProject = null;
    var savedProjects = [];
    var projectEditMode = 'create';
    var ductStaticSegments = [];
    var ductStaticIdSeq = 1;
    var condensateZones = [];
    var sanitaryZones = [];
    var wsfuZones = [];
    var ductZones = [];
    var condensateZoneIdSeq = 0;
    var condensateRowIdSeq = 0;
    var sanitaryZoneIdSeq = 0;
    var wsfuZoneIdSeq = 0;
    var ductZoneIdSeq = 0;
    var ductSegmentIdSeq = 0;
    var ventilationZones = [];
    var ventilationZoneIdSeq = 0;
    var ventilationEditingZoneId = '';
    var ventFixtureRows = [];
    var ventFixtureIdSeq = 0;
    var wsfuFixtureRows = [];
    var wsfuFixtureIdSeq = 1;
    var wsfuFixtureCatalog = [
      { key: 'BATHTUB_SHOWER_FILL', label: 'Bathtub / Shower (fill)', privateWsfu: 4.0, publicWsfu: 4.0, assemblyWsfu: null },
      { key: 'BATHTUB_FILL_VALVE_3_4', label: 'Bathtub Fill Valve (3/4")', privateWsfu: 10.0, publicWsfu: 10.0, assemblyWsfu: null },
      { key: 'BIDET', label: 'Bidet', privateWsfu: 1.0, publicWsfu: null, assemblyWsfu: null },
      { key: 'CLOTHES_WASHER', label: 'Clothes Washer', privateWsfu: 4.0, publicWsfu: 4.0, assemblyWsfu: null },
      { key: 'DENTAL_UNIT', label: 'Dental Unit', privateWsfu: null, publicWsfu: 1.0, assemblyWsfu: null },
      { key: 'DISHWASHER', label: 'Dishwasher', privateWsfu: 1.5, publicWsfu: 1.5, assemblyWsfu: null },
      { key: 'DRINKING_FOUNTAIN', label: 'Drinking Fountain', privateWsfu: 0.5, publicWsfu: 0.5, assemblyWsfu: 0.75 },
      { key: 'HOSE_BIBB', label: 'Hose Bibb', privateWsfu: 2.5, publicWsfu: 2.5, assemblyWsfu: null },
      { key: 'HOSE_BIBB_ADDITIONAL', label: 'Hose Bibb (additional)', privateWsfu: 1.0, publicWsfu: 1.0, assemblyWsfu: null },
      { key: 'LAVATORY', label: 'Lavatory', privateWsfu: 1.0, publicWsfu: 1.0, assemblyWsfu: 1.0 },
      { key: 'LAWN_SPRINKLER', label: 'Lawn Sprinkler', privateWsfu: 1.0, publicWsfu: 1.0, assemblyWsfu: null },
      { key: 'MOBILE_HOME', label: 'Mobile Home', privateWsfu: 6.0, publicWsfu: null, assemblyWsfu: null },
      { key: 'SINK_BAR', label: 'Sink - Bar', privateWsfu: 1.0, publicWsfu: 2.0, assemblyWsfu: null },
      { key: 'SINK_CLINICAL', label: 'Sink - Clinical', privateWsfu: null, publicWsfu: 3.0, assemblyWsfu: null },
      { key: 'KITCHEN_SINK', label: 'Sink - Kitchen', privateWsfu: 1.5, publicWsfu: 1.5, assemblyWsfu: null },
      { key: 'SINK_LAUNDRY', label: 'Sink - Laundry', privateWsfu: 1.5, publicWsfu: 1.5, assemblyWsfu: null },
      { key: 'SERVICE_MOP_BASIN', label: 'Service / Mop Basin', privateWsfu: 1.5, publicWsfu: 3.0, assemblyWsfu: null },
      { key: 'WASH_FOUNTAIN', label: 'Wash Fountain', privateWsfu: null, publicWsfu: 2.0, assemblyWsfu: null },
      { key: 'SHOWER_SINGLE', label: 'Shower (per head)', privateWsfu: 2.0, publicWsfu: 2.0, assemblyWsfu: null },
      { key: 'URINAL_1_0_GPF', label: 'Urinal (1.0 GPF)', privateWsfu: null, publicWsfu: null, assemblyWsfu: null },
      { key: 'URINAL_TANK', label: 'Urinal (tank)', privateWsfu: 2.0, publicWsfu: 2.0, assemblyWsfu: 3.0 },
      { key: 'URINAL_WITH_FLUSH_ACTION', label: 'Urinal (with flush action)', privateWsfu: 1.0, publicWsfu: 1.0, assemblyWsfu: 1.0 },
      { key: 'WATER_CLOSET_GRAVITY', label: 'Water Closet (1.6 GPF Gravity Tank)', privateWsfu: 2.5, publicWsfu: 2.5, assemblyWsfu: 3.5 },
      { key: 'WATER_CLOSET_FLUSHOMETER', label: 'Water Closet (1.6 GPF Flushometer)', privateWsfu: 2.5, publicWsfu: 2.5, assemblyWsfu: 3.5 },
      { key: 'WATER_CLOSET_GT_1_6_GRAVITY', label: 'Water Closet (>1.6 GPF Gravity Tank)', privateWsfu: 3.0, publicWsfu: 5.5, assemblyWsfu: 7.0 },
      { key: 'WATER_CLOSET_GT_1_6_FLUSHOMETER', label: 'Water Closet (>1.6 GPF Flushometer)', privateWsfu: null, publicWsfu: null, assemblyWsfu: null }
    ];
    var ventFixtureCatalog = [
      { key: 'BATHTUB_BATH_SHOWER', label: 'Bathtub / Shower', privateDfu: 2.0, publicDfu: 2.0, assemblyDfu: null },
      { key: 'BIDET', label: 'Bidet', privateDfu: 1.0, publicDfu: null, assemblyDfu: null },
      { key: 'BIDET_ALT', label: 'Bidet (alt)', privateDfu: 2.0, publicDfu: null, assemblyDfu: null },
      { key: 'CLOTHES_WASHER', label: 'Clothes Washer', privateDfu: 3.0, publicDfu: 3.0, assemblyDfu: 3.0 },
      { key: 'DENTAL_UNIT', label: 'Dental Unit', privateDfu: null, publicDfu: 1.0, assemblyDfu: 1.0 },
      { key: 'DISHWASHER', label: 'Dishwasher', privateDfu: 2.0, publicDfu: 2.0, assemblyDfu: 2.0 },
      { key: 'DRINKING_FOUNTAIN', label: 'Drinking Fountain', privateDfu: 0.5, publicDfu: 0.5, assemblyDfu: 1.0 },
      { key: 'FOOD_WASTE_DISPOSER_COMMERCIAL', label: 'Food Waste Disposer (commercial)', privateDfu: null, publicDfu: 3.0, assemblyDfu: 3.0 },
      { key: 'FLOOR_DRAIN', label: 'Floor Drain', privateDfu: 2.0, publicDfu: 2.0, assemblyDfu: 2.0 },
      { key: 'SHOWER_SINGLE', label: 'Shower', privateDfu: 2.0, publicDfu: 2.0, assemblyDfu: 2.0 },
      { key: 'SHOWER_MULTI_HEAD_ADDITIONAL', label: 'Shower (multi-head additional)', privateDfu: 1.0, publicDfu: 1.0, assemblyDfu: 1.0 },
      { key: 'LAVATORY', label: 'Lavatory', privateDfu: 1.0, publicDfu: 1.0, assemblyDfu: 1.0 },
      { key: 'LAVATORIES_SET', label: 'Lavatories (set)', privateDfu: 2.0, publicDfu: 2.0, assemblyDfu: 2.0 },
      { key: 'WASH_FOUNTAIN', label: 'Wash Fountain', privateDfu: null, publicDfu: 2.0, assemblyDfu: 2.0 },
      { key: 'MOBILE_HOME_TRAP', label: 'Mobile Home Trap', privateDfu: 6.0, publicDfu: null, assemblyDfu: null },
      { key: 'SINK_BAR', label: 'Sink - Bar', privateDfu: 1.0, publicDfu: null, assemblyDfu: null },
      { key: 'SINK_CLINICAL', label: 'Sink - Clinical', privateDfu: null, publicDfu: 6.0, assemblyDfu: 6.0 },
      { key: 'SINK_COMMERCIAL', label: 'Sink - Commercial', privateDfu: null, publicDfu: 3.0, assemblyDfu: 3.0 },
      { key: 'KITCHEN_SINK', label: 'Sink - Kitchen', privateDfu: 2.0, publicDfu: 2.0, assemblyDfu: null },
      { key: 'SINK_LAUNDRY', label: 'Sink - Laundry', privateDfu: 2.0, publicDfu: 2.0, assemblyDfu: 2.0 },
      { key: 'SERVICE_MOP_BASIN', label: 'Service Sink', privateDfu: null, publicDfu: 3.0, assemblyDfu: 3.0 },
      { key: 'URINAL_INTEGRAL_TRAP_1_0_GPF', label: 'Urinal (integral trap 1.0 GPF)', privateDfu: 2.0, publicDfu: 2.0, assemblyDfu: 5.0 },
      { key: 'URINAL_GT_1_0_GPF', label: 'Urinal (>1.0 GPF)', privateDfu: 2.0, publicDfu: 2.0, assemblyDfu: 6.0 },
      { key: 'URINAL_EXPOSED_TRAP', label: 'Urinal (exposed trap)', privateDfu: 2.0, publicDfu: 2.0, assemblyDfu: 5.0 },
      { key: 'WATER_CLOSET_GRAVITY', label: 'Water Closet (1.6 GPF Gravity Tank)', privateDfu: 3.0, publicDfu: 4.0, assemblyDfu: 6.0 },
      { key: 'WATER_CLOSET_FLUSHOMETER', label: 'Water Closet (1.6 GPF Flushometer)', privateDfu: 3.0, publicDfu: 4.0, assemblyDfu: 6.0 },
      { key: 'WATER_CLOSET_GT_1_6_GRAVITY', label: 'Water Closet (>1.6 GPF Gravity Tank)', privateDfu: 4.0, publicDfu: 6.0, assemblyDfu: 8.0 },
      { key: 'WATER_CLOSET_GT_1_6_FLUSHOMETER', label: 'Water Closet (>1.6 GPF Flushometer)', privateDfu: 4.0, publicDfu: 6.0, assemblyDfu: 8.0 }
    ];
    var fixtureUnitCatalog = [
      { key: 'BATHTUB_BATH_SHOWER', label: 'Bathtub / Shower', water: { private: 4.0, public: 4.0, assembly: null }, waste: { private: 2.0, public: 2.0, assembly: null } },
      { key: 'BATHTUB_FILL_VALVE_3_4', label: 'Bathtub Fill Valve (3/4")', water: { private: 10.0, public: 10.0, assembly: null }, waste: { private: null, public: null, assembly: null } },
      { key: 'BIDET', label: 'Bidet', water: { private: 1.0, public: null, assembly: null }, waste: { private: 1.0, public: null, assembly: null } },
      { key: 'BIDET_ALT', label: 'Bidet (alt)', water: { private: null, public: null, assembly: null }, waste: { private: 2.0, public: null, assembly: null } },
      { key: 'CLOTHES_WASHER', label: 'Clothes Washer', water: { private: 4.0, public: 4.0, assembly: null }, waste: { private: 3.0, public: 3.0, assembly: 3.0 } },
      { key: 'DENTAL_UNIT', label: 'Dental Unit', water: { private: null, public: 1.0, assembly: null }, waste: { private: null, public: 1.0, assembly: 1.0 } },
      { key: 'DISHWASHER', label: 'Dishwasher', water: { private: 1.5, public: 1.5, assembly: null }, waste: { private: 2.0, public: 2.0, assembly: 2.0 } },
      { key: 'DRINKING_FOUNTAIN', label: 'Drinking Fountain', water: { private: 0.5, public: 0.5, assembly: 0.75 }, waste: { private: 0.5, public: 0.5, assembly: 1.0 } },
      { key: 'FOOD_WASTE_DISPOSER_COMMERCIAL', label: 'Food Waste Disposer (commercial)', water: { private: null, public: null, assembly: null }, waste: { private: null, public: 3.0, assembly: 3.0 } },
      { key: 'FLOOR_DRAIN', label: 'Floor Drain', water: { private: null, public: null, assembly: null }, waste: { private: 2.0, public: 2.0, assembly: 2.0 } },
      { key: 'HOSE_BIBB', label: 'Hose Bibb', water: { private: 2.5, public: 2.5, assembly: null }, waste: { private: null, public: null, assembly: null } },
      { key: 'HOSE_BIBB_ADDITIONAL', label: 'Hose Bibb (additional)', water: { private: 1.0, public: 1.0, assembly: null }, waste: { private: null, public: null, assembly: null } },
      { key: 'LAVATORY', label: 'Lavatory', water: { private: 1.0, public: 1.0, assembly: 1.0 }, waste: { private: 1.0, public: 1.0, assembly: 1.0 } },
      { key: 'LAVATORIES_SET', label: 'Lavatories (set)', water: { private: null, public: null, assembly: null }, waste: { private: 2.0, public: 2.0, assembly: 2.0 } },
      { key: 'LAWN_SPRINKLER', label: 'Lawn Sprinkler', water: { private: 1.0, public: 1.0, assembly: null }, waste: { private: null, public: null, assembly: null } },
      { key: 'MOBILE_HOME', label: 'Mobile Home', water: { private: 6.0, public: null, assembly: null }, waste: { private: null, public: null, assembly: null } },
      { key: 'MOBILE_HOME_TRAP', label: 'Mobile Home Trap', water: { private: null, public: null, assembly: null }, waste: { private: 6.0, public: null, assembly: null } },
      { key: 'SINK_BAR', label: 'Sink - Bar', water: { private: 1.0, public: 2.0, assembly: null }, waste: { private: 1.0, public: null, assembly: null } },
      { key: 'SINK_CLINICAL', label: 'Sink - Clinical', water: { private: null, public: 3.0, assembly: null }, waste: { private: null, public: 6.0, assembly: 6.0 } },
      { key: 'SINK_COMMERCIAL', label: 'Sink - Commercial', water: { private: null, public: null, assembly: null }, waste: { private: null, public: 3.0, assembly: 3.0 } },
      { key: 'KITCHEN_SINK', label: 'Sink - Kitchen', water: { private: 1.5, public: 1.5, assembly: null }, waste: { private: 2.0, public: 2.0, assembly: null } },
      { key: 'SINK_LAUNDRY', label: 'Sink - Laundry', water: { private: 1.5, public: 1.5, assembly: null }, waste: { private: 2.0, public: 2.0, assembly: 2.0 } },
      { key: 'SERVICE_MOP_BASIN', label: 'Service / Mop Basin', water: { private: 1.5, public: 3.0, assembly: null }, waste: { private: null, public: 3.0, assembly: 3.0 } },
      { key: 'SHOWER_SINGLE', label: 'Shower', water: { private: 2.0, public: 2.0, assembly: null }, waste: { private: 2.0, public: 2.0, assembly: 2.0 } },
      { key: 'SHOWER_MULTI_HEAD_ADDITIONAL', label: 'Shower (multi-head additional)', water: { private: null, public: null, assembly: null }, waste: { private: 1.0, public: 1.0, assembly: 1.0 } },
      { key: 'URINAL_1_0_GPF', label: 'Urinal (1.0 GPF)', water: { private: null, public: null, assembly: null }, waste: { private: null, public: null, assembly: null } },
      { key: 'URINAL_TANK', label: 'Urinal (tank)', water: { private: 2.0, public: 2.0, assembly: 3.0 }, waste: { private: null, public: null, assembly: null } },
      { key: 'URINAL_WITH_FLUSH_ACTION', label: 'Urinal (with flush action)', water: { private: 1.0, public: 1.0, assembly: 1.0 }, waste: { private: null, public: null, assembly: null } },
      { key: 'URINAL_INTEGRAL_TRAP_1_0_GPF', label: 'Urinal (integral trap 1.0 GPF)', water: { private: null, public: null, assembly: null }, waste: { private: 2.0, public: 2.0, assembly: 5.0 } },
      { key: 'URINAL_GT_1_0_GPF', label: 'Urinal (>1.0 GPF)', water: { private: null, public: null, assembly: null }, waste: { private: 2.0, public: 2.0, assembly: 6.0 } },
      { key: 'URINAL_EXPOSED_TRAP', label: 'Urinal (exposed trap)', water: { private: null, public: null, assembly: null }, waste: { private: 2.0, public: 2.0, assembly: 5.0 } },
      { key: 'WASH_FOUNTAIN', label: 'Wash Fountain', water: { private: null, public: 2.0, assembly: null }, waste: { private: null, public: 2.0, assembly: 2.0 } },
      { key: 'WATER_CLOSET_GRAVITY', label: 'Water Closet (1.6 GPF Gravity Tank)', water: { private: 2.5, public: 2.5, assembly: 3.5 }, waste: { private: 3.0, public: 4.0, assembly: 6.0 } },
      { key: 'WATER_CLOSET_FLUSHOMETER', label: 'Water Closet (1.6 GPF Flushometer)', water: { private: 2.5, public: 2.5, assembly: 3.5 }, waste: { private: 3.0, public: 4.0, assembly: 6.0 } },
      { key: 'WATER_CLOSET_GT_1_6_GRAVITY', label: 'Water Closet (>1.6 GPF Gravity Tank)', water: { private: 3.0, public: 5.5, assembly: 7.0 }, waste: { private: 4.0, public: 6.0, assembly: 8.0 } },
      { key: 'WATER_CLOSET_GT_1_6_FLUSHOMETER', label: 'Water Closet (>1.6 GPF Flushometer)', water: { private: null, public: null, assembly: null }, waste: { private: 4.0, public: 6.0, assembly: 8.0 } }
    ];
    function makeVentOcc(group, key, label, rp, ra, density, airClass, aliases, sortOrder) {
      return {
        category_group: group,
        key: key,
        occupancy_category: label,
        Rp: rp,
        Ra: ra,
        default_occupant_density: density,
        air_class: airClass,
        searchable_aliases: aliases || [],
        sort_order: sortOrder
      };
    }
    var ventilationOccupancyCatalog = [
      makeVentOcc('ANIMAL FACILITIES', 'ANIMAL_EXAM_ROOM_VET', 'Animal exam room (veterinary office)', 10.0, 0.12, 20.0, 2, ['animal', 'vet', 'veterinary', 'exam'], 10),
      makeVentOcc('ANIMAL FACILITIES', 'ANIMAL_IMAGING', 'Animal imaging (MRI/CT/PET)', 10.0, 0.18, 20.0, 3, ['animal', 'imaging', 'mri', 'ct', 'pet'], 20),
      makeVentOcc('ANIMAL FACILITIES', 'ANIMAL_OPERATING_ROOMS', 'Animal operating rooms', 10.0, 0.18, 20.0, 3, ['animal', 'operating', 'surgery'], 30),
      makeVentOcc('ANIMAL FACILITIES', 'ANIMAL_POSTOP_RECOVERY', 'Animal postoperative recovery room', 10.0, 0.18, 20.0, 3, ['animal', 'postop', 'recovery'], 40),
      makeVentOcc('ANIMAL FACILITIES', 'ANIMAL_PREPARATION', 'Animal preparation rooms', 10.0, 0.18, 20.0, 3, ['animal', 'prep'], 50),
      makeVentOcc('ANIMAL FACILITIES', 'ANIMAL_PROCEDURE', 'Animal procedure room', 10.0, 0.18, 20.0, 3, ['animal', 'procedure'], 60),
      makeVentOcc('ANIMAL FACILITIES', 'ANIMAL_SURGERY_SCRUB', 'Animal surgery scrub', 10.0, 0.18, 20.0, 3, ['animal', 'scrub'], 70),
      makeVentOcc('ANIMAL FACILITIES', 'LARGE_ANIMAL_HOLDING', 'Large-animal holding room', 10.0, 0.18, 20.0, 3, ['animal', 'holding'], 80),
      makeVentOcc('ANIMAL FACILITIES', 'NECROPSY', 'Necropsy', 10.0, 0.18, 20.0, 3, ['animal', 'necropsy'], 90),
      makeVentOcc('ANIMAL FACILITIES', 'SMALL_ANIMAL_CAGE_STATIC', 'Small-animal-cage room (static cages)', 10.0, 0.18, 20.0, 3, ['animal', 'cage', 'static'], 100),
      makeVentOcc('ANIMAL FACILITIES', 'SMALL_ANIMAL_CAGE_VENT', 'Small-animal-cage room (ventilated cages)', 10.0, 0.18, 20.0, 3, ['animal', 'cage', 'ventilated'], 110),

      makeVentOcc('CORRECTIONAL FACILITIES', 'BOOKING_WAITING', 'Booking/waiting', 7.5, 0.06, 50.0, 2, ['booking', 'waiting'], 10),
      makeVentOcc('CORRECTIONAL FACILITIES', 'CELL', 'Cell', 5.0, 0.12, 25.0, 2, ['cell'], 20),
      makeVentOcc('CORRECTIONAL FACILITIES', 'DAY_ROOM', 'Day room', 5.0, 0.06, 30.0, 1, ['day room'], 30),
      makeVentOcc('CORRECTIONAL FACILITIES', 'GUARD_STATIONS', 'Guard stations', 5.0, 0.06, 15.0, 1, ['guard'], 40),

      makeVentOcc('EDUCATIONAL FACILITIES', 'ART_CLASSROOM', 'Art classroom', 10.0, 0.18, 20.0, 2, ['art', 'classroom'], 10),
      makeVentOcc('EDUCATIONAL FACILITIES', 'CLASSROOM_AGE_5_8', 'Classrooms (ages 5 to 8)', 10.0, 0.12, 25.0, 1, ['classroom', 'school'], 20),
      makeVentOcc('EDUCATIONAL FACILITIES', 'CLASSROOM_AGE_9_PLUS', 'Classrooms (age 9 plus)', 10.0, 0.12, 35.0, 1, ['classroom', 'school'], 30),
      makeVentOcc('EDUCATIONAL FACILITIES', 'COMPUTER_LAB', 'Computer lab', 10.0, 0.12, 25.0, 1, ['computer', 'lab'], 40),
      makeVentOcc('EDUCATIONAL FACILITIES', 'DAYCARE_SICKROOM', 'Daycare sickroom', 10.0, 0.18, 25.0, 3, ['daycare', 'sickroom'], 50),
      makeVentOcc('EDUCATIONAL FACILITIES', 'DAYCARE_THROUGH_4', 'Daycare (through age 4)', 10.0, 0.18, 25.0, 2, ['daycare', 'preschool'], 60),
      makeVentOcc('EDUCATIONAL FACILITIES', 'LECTURE_CLASSROOM', 'Lecture classroom', 7.5, 0.06, 65.0, 1, ['lecture', 'classroom'], 70),
      makeVentOcc('EDUCATIONAL FACILITIES', 'LECTURE_HALL', 'Lecture hall (fixed seats)', 7.5, 0.06, 150.0, 1, ['lecture hall'], 80),
      makeVentOcc('EDUCATIONAL FACILITIES', 'LIBRARIES_EDU', 'Libraries', 5.0, 0.12, 10.0, 1, ['library'], 90),
      makeVentOcc('EDUCATIONAL FACILITIES', 'MEDIA_CENTER', 'Media center', 10.0, 0.12, 25.0, 1, ['media center'], 100),
      makeVentOcc('EDUCATIONAL FACILITIES', 'MULTI_USE_ASSEMBLY', 'Multi-use assembly', 7.5, 0.06, 100.0, 1, ['assembly'], 110),
      makeVentOcc('EDUCATIONAL FACILITIES', 'MUSIC_THEATER_DANCE', 'Music/theater/dance', 10.0, 0.06, 35.0, 1, ['music', 'theater', 'dance'], 120),
      makeVentOcc('EDUCATIONAL FACILITIES', 'SCIENCE_LABS', 'Science laboratories', 10.0, 0.18, 25.0, 2, ['science', 'lab'], 130),
      makeVentOcc('EDUCATIONAL FACILITIES', 'UNIVERSITY_COLLEGE_LABS', 'University/college laboratories', 10.0, 0.18, 25.0, 2, ['university', 'college', 'lab'], 140),
      makeVentOcc('EDUCATIONAL FACILITIES', 'WOOD_METAL_SHOP', 'Wood/metal shop', 10.0, 0.18, 20.0, 2, ['shop', 'wood', 'metal'], 150),

      makeVentOcc('FOOD AND BEVERAGE SERVICE', 'BARS_COCKTAIL_LOUNGES', 'Bars, cocktail lounges', 7.5, 0.18, 100.0, 2, ['bar', 'cocktail', 'lounge'], 10),
      makeVentOcc('FOOD AND BEVERAGE SERVICE', 'CAFETERIA_FAST_FOOD', 'Cafeteria/fast-food dining', 7.5, 0.18, 100.0, 2, ['cafeteria', 'fast food'], 20),
      makeVentOcc('FOOD AND BEVERAGE SERVICE', 'KITCHEN_COOKING', 'Kitchen (cooking)', 7.5, 0.12, 20.0, 2, ['kitchen', 'cooking'], 30),
      makeVentOcc('FOOD AND BEVERAGE SERVICE', 'RESTAURANT_DINING', 'Restaurant dining rooms', 7.5, 0.18, 70.0, 2, ['restaurant', 'dining'], 40),

      makeVentOcc('GENERAL', 'BREAK_ROOMS_GENERAL', 'Break rooms', 5.0, 0.06, 25.0, 1, ['break room'], 10),
      makeVentOcc('GENERAL', 'COFFEE_STATIONS', 'Coffee stations', 5.0, 0.06, 20.0, 1, ['coffee'], 20),
      makeVentOcc('GENERAL', 'CONFERENCE_MEETING', 'Conference/meeting', 5.0, 0.06, 50.0, 1, ['conference', 'meeting room'], 30),
      makeVentOcc('GENERAL', 'CORRIDORS', 'Corridors', 0.0, 0.06, 0.0, 1, ['corridor', 'hallway'], 40),
      makeVentOcc('GENERAL', 'OCCUP_STORAGE_LIQUIDS_GELS', 'Occupiable storage rooms for liquids or gels', 5.0, 0.12, 2.0, 2, ['storage', 'liquids', 'gels'], 50),

      makeVentOcc('HOTELS / MOTELS / RESORTS / DORMITORIES', 'BARRACKS_SLEEP', 'Barracks sleeping areas', 5.0, 0.06, 20.0, 1, ['barracks', 'sleeping'], 10),
      makeVentOcc('HOTELS / MOTELS / RESORTS / DORMITORIES', 'BEDROOM_LIVING', 'Bedroom/living room', 5.0, 0.06, 10.0, 1, ['bedroom', 'living'], 20),
      makeVentOcc('HOTELS / MOTELS / RESORTS / DORMITORIES', 'LAUNDRY_CENTRAL', 'Laundry rooms, central', 5.0, 0.12, 10.0, 2, ['laundry', 'central'], 30),
      makeVentOcc('HOTELS / MOTELS / RESORTS / DORMITORIES', 'LAUNDRY_DWELLING', 'Laundry rooms within dwelling units', 5.0, 0.12, 10.0, 1, ['laundry', 'dwelling'], 40),
      makeVentOcc('HOTELS / MOTELS / RESORTS / DORMITORIES', 'LOBBIES_PREFUNCTION', 'Lobbies/prefunction', 7.5, 0.06, 30.0, 1, ['lobby', 'prefunction'], 50),
      makeVentOcc('HOTELS / MOTELS / RESORTS / DORMITORIES', 'MULTIPURPOSE_ASSEMBLY', 'Multipurpose assembly', 5.0, 0.06, 120.0, 1, ['multipurpose', 'assembly'], 60),

      makeVentOcc('MISCELLANEOUS SPACES', 'BANKS_LOBBIES', 'Banks or bank lobbies', 7.5, 0.06, 15.0, 1, ['bank', 'lobby'], 10),
      makeVentOcc('MISCELLANEOUS SPACES', 'BANK_VAULTS', 'Bank vaults/safe deposit', 5.0, 0.06, 5.0, 2, ['vault', 'safe deposit'], 20),
      makeVentOcc('MISCELLANEOUS SPACES', 'COMPUTER_NOT_PRINTING', 'Computer (not printing)', 5.0, 0.06, 4.0, 1, ['computer room'], 30),
      makeVentOcc('MISCELLANEOUS SPACES', 'FREEZER_REFRIG_SPACES', 'Freezer and refrigerated spaces (<50°F)', 10.0, 0.0, 0.0, 2, ['freezer', 'refrigerated'], 40),
      makeVentOcc('MISCELLANEOUS SPACES', 'MANUFACTURING_NON_HAZ', 'Manufacturing where hazardous materials are not used', 10.0, 0.18, 7.0, 2, ['manufacturing'], 50),
      makeVentOcc('MISCELLANEOUS SPACES', 'MANUFACTURING_HAZ', 'Manufacturing where hazardous materials are used (excluding heavy industrial / chemical processes)', 10.0, 0.18, 7.0, 3, ['manufacturing', 'hazardous'], 60),
      makeVentOcc('MISCELLANEOUS SPACES', 'PHARMACY_PREP', 'Pharmacy (prep. area)', 5.0, 0.18, 10.0, 2, ['pharmacy'], 70),
      makeVentOcc('MISCELLANEOUS SPACES', 'PHOTO_STUDIOS', 'Photo studios', 5.0, 0.12, 10.0, 1, ['photo', 'studio'], 80),
      makeVentOcc('MISCELLANEOUS SPACES', 'SHIPPING_RECEIVING', 'Shipping/receiving', 10.0, 0.12, 2.0, 2, ['shipping', 'receiving'], 90),
      makeVentOcc('MISCELLANEOUS SPACES', 'SORTING_PACKING_LIGHT_ASSEMBLY', 'Sorting, packing, light assembly', 7.5, 0.12, 7.0, 2, ['sorting', 'packing', 'assembly'], 100),
      makeVentOcc('MISCELLANEOUS SPACES', 'TELEPHONE_CLOSETS', 'Telephone closets', 0.0, 0.0, 0.0, 1, ['telephone closet', 'data closet'], 110),
      makeVentOcc('MISCELLANEOUS SPACES', 'TRANSPORT_WAITING', 'Transportation waiting', 7.5, 0.06, 100.0, 1, ['transportation waiting', 'waiting area'], 120),
      makeVentOcc('MISCELLANEOUS SPACES', 'WAREHOUSES', 'Warehouses', 10.0, 0.06, 0.0, 2, ['warehouse'], 130),

      makeVentOcc('OFFICE BUILDINGS', 'BREAK_ROOMS_OFFICE', 'Break Rooms', 5.0, 0.12, 50.0, 1, ['office break room'], 10),
      makeVentOcc('OFFICE BUILDINGS', 'MAIN_ENTRY_LOBBIES', 'Main entry lobbies', 5.0, 0.06, 10.0, 1, ['lobby', 'entry'], 20),
      makeVentOcc('OFFICE BUILDINGS', 'OCCUP_STORAGE_DRY', 'Occupiable storage rooms for dry materials', 5.0, 0.06, 2.0, 1, ['storage', 'dry'], 30),
      makeVentOcc('OFFICE BUILDINGS', 'OFFICE_SPACE', 'Office space', 5.0, 0.06, 5.0, 1, ['office'], 40),
      makeVentOcc('OFFICE BUILDINGS', 'RECEPTION_AREAS', 'Reception areas', 5.0, 0.06, 30.0, 1, ['reception'], 50),
      makeVentOcc('OFFICE BUILDINGS', 'TELEPHONE_DATA_ENTRY', 'Telephone/data entry', 5.0, 0.06, 60.0, 1, ['data entry', 'telephone'], 60),

      makeVentOcc('OUTPATIENT HEALTH CARE FACILITIES', 'BIRTHING_ROOM', 'Birthing room', 10.0, 0.18, 15.0, 2, ['birthing'], 10),
      makeVentOcc('OUTPATIENT HEALTH CARE FACILITIES', 'CLASS1_IMAGING', 'Class 1 imaging rooms', 5.0, 0.12, 5.0, 1, ['imaging'], 20),
      makeVentOcc('OUTPATIENT HEALTH CARE FACILITIES', 'DENTAL_OPERATORY', 'Dental operatory', 10.0, 0.18, 20.0, 1, ['dental'], 30),
      makeVentOcc('OUTPATIENT HEALTH CARE FACILITIES', 'GENERAL_EXAM_ROOM', 'General examination room', 7.5, 0.12, 20.0, 1, ['exam room'], 40),
      makeVentOcc('OUTPATIENT HEALTH CARE FACILITIES', 'OTHER_DENTAL_TREATMENT', 'Other dental treatment areas', 5.0, 0.06, 5.0, 1, ['dental treatment'], 50),
      makeVentOcc('OUTPATIENT HEALTH CARE FACILITIES', 'PT_EXERCISE', 'Physical therapy exercise area', 20.0, 0.18, 7.0, 2, ['physical therapy', 'exercise'], 60),
      makeVentOcc('OUTPATIENT HEALTH CARE FACILITIES', 'PT_INDIVIDUAL_ROOM', 'Physical therapy individual room', 10.0, 0.06, 20.0, 1, ['physical therapy'], 70),
      makeVentOcc('OUTPATIENT HEALTH CARE FACILITIES', 'PT_POOL_AREA', 'Physical therapeutic pool area', 0.0, 0.48, 0.0, 2, ['pool area'], 80),
      makeVentOcc('OUTPATIENT HEALTH CARE FACILITIES', 'PROSTHETICS_ORTHOTICS', 'Prosthetics and orthotics room', 10.0, 0.18, 20.0, 1, ['prosthetics', 'orthotics'], 90),
      makeVentOcc('OUTPATIENT HEALTH CARE FACILITIES', 'PSYCH_CONSULT', 'Psychiatric consultation room', 5.0, 0.06, 20.0, 1, ['psychiatric', 'consultation'], 100),
      makeVentOcc('OUTPATIENT HEALTH CARE FACILITIES', 'PSYCH_EXAM_ROOM', 'Psychiatric examination room', 5.0, 0.06, 20.0, 1, ['psychiatric', 'exam'], 110),
      makeVentOcc('OUTPATIENT HEALTH CARE FACILITIES', 'PSYCH_GROUP_ROOM', 'Psychiatric group room', 5.0, 0.06, 50.0, 1, ['psychiatric', 'group'], 120),
      makeVentOcc('OUTPATIENT HEALTH CARE FACILITIES', 'PSYCH_SECLUSION', 'Psychiatric seclusion room', 10.0, 0.06, 5.0, 1, ['seclusion'], 130),
      makeVentOcc('OUTPATIENT HEALTH CARE FACILITIES', 'SPEECH_THERAPY', 'Speech therapy room', 5.0, 0.06, 20.0, 1, ['speech therapy'], 140),
      makeVentOcc('OUTPATIENT HEALTH CARE FACILITIES', 'URGENT_CARE_EXAM', 'Urgent care examination room', 7.5, 0.12, 20.0, 1, ['urgent care'], 150),
      makeVentOcc('OUTPATIENT HEALTH CARE FACILITIES', 'URGENT_CARE_OBSERVATION', 'Urgent care observation room', 5.0, 0.06, 20.0, 1, ['urgent care'], 160),
      makeVentOcc('OUTPATIENT HEALTH CARE FACILITIES', 'URGENT_CARE_TREATMENT', 'Urgent care treatment room', 7.5, 0.18, 20.0, 1, ['urgent care'], 170),
      makeVentOcc('OUTPATIENT HEALTH CARE FACILITIES', 'URGENT_CARE_TRIAGE', 'Urgent care triage room', 10.0, 0.18, 20.0, 1, ['urgent care', 'triage'], 180),

      makeVentOcc('PUBLIC ASSEMBLY SPACES', 'AUDITORIUM_SEATING', 'Auditorium seating area', 5.0, 0.06, 150.0, 1, ['auditorium'], 10),
      makeVentOcc('PUBLIC ASSEMBLY SPACES', 'COURTROOMS', 'Courtrooms', 5.0, 0.06, 70.0, 1, ['courtroom'], 20),
      makeVentOcc('PUBLIC ASSEMBLY SPACES', 'LEGISLATIVE_CHAMBERS', 'Legislative chambers', 5.0, 0.06, 50.0, 1, ['legislative'], 30),
      makeVentOcc('PUBLIC ASSEMBLY SPACES', 'LIBRARIES_PUBLIC', 'Libraries', 5.0, 0.12, 10.0, 1, ['library'], 40),
      makeVentOcc('PUBLIC ASSEMBLY SPACES', 'LOBBIES_PUBLIC', 'Lobbies', 5.0, 0.06, 150.0, 1, ['lobby'], 50),
      makeVentOcc('PUBLIC ASSEMBLY SPACES', 'MUSEUM_CHILDRENS', 'Museums (children’s)', 7.5, 0.12, 40.0, 1, ['museum'], 60),
      makeVentOcc('PUBLIC ASSEMBLY SPACES', 'MUSEUM_GALLERY', 'Museums/galleries', 7.5, 0.06, 40.0, 1, ['museum', 'gallery'], 70),
      makeVentOcc('PUBLIC ASSEMBLY SPACES', 'RELIGIOUS_WORSHIP', 'Places of religious worship', 5.0, 0.06, 120.0, 1, ['religious', 'worship'], 80),

      makeVentOcc('RETAIL', 'RETAIL_SALES', 'Sales (except as below)', 7.5, 0.12, 15.0, 2, ['retail', 'sales'], 10),
      makeVentOcc('RETAIL', 'BARBER_SHOP', 'Barber shop', 7.5, 0.06, 25.0, 2, ['barber'], 20),
      makeVentOcc('RETAIL', 'BEAUTY_NAIL_SALONS', 'Beauty and nail salons', 20.0, 0.12, 25.0, 2, ['beauty', 'nail'], 30),
      makeVentOcc('RETAIL', 'COIN_LAUNDRIES', 'Coin-operated laundries', 7.5, 0.12, 20.0, 2, ['coin laundry'], 40),
      makeVentOcc('RETAIL', 'MALL_COMMON_AREAS', 'Mall common areas', 7.5, 0.06, 40.0, 1, ['mall'], 50),
      makeVentOcc('RETAIL', 'PET_SHOPS_ANIMAL_AREAS', 'Pet shops (animal areas)', 7.5, 0.18, 10.0, 2, ['pet shop'], 60),
      makeVentOcc('RETAIL', 'SUPERMARKET', 'Supermarket', 7.5, 0.06, 8.0, 1, ['supermarket', 'grocery'], 70),

      makeVentOcc('SPORTS AND ENTERTAINMENT', 'BOWLING_SEATING', 'Bowling alley (seating)', 10.0, 0.12, 40.0, 1, ['bowling'], 10),
      makeVentOcc('SPORTS AND ENTERTAINMENT', 'DISCO_DANCE', 'Disco/dance floors', 20.0, 0.06, 100.0, 2, ['dance floor', 'disco'], 20),
      makeVentOcc('SPORTS AND ENTERTAINMENT', 'GAMBLING_CASINOS', 'Gambling casinos', 7.5, 0.18, 120.0, 1, ['casino'], 30),
      makeVentOcc('SPORTS AND ENTERTAINMENT', 'GAME_ARCADES', 'Game arcades', 7.5, 0.18, 20.0, 1, ['arcade'], 40),
      makeVentOcc('SPORTS AND ENTERTAINMENT', 'GYM_PLAY', 'Gym, sports arena (play area)', 20.0, 0.18, 7.0, 2, ['gym', 'sports arena'], 50),
      makeVentOcc('SPORTS AND ENTERTAINMENT', 'HEALTH_CLUB_AEROBICS', 'Health club/aerobics room', 20.0, 0.06, 40.0, 2, ['health club', 'aerobics'], 60),
      makeVentOcc('SPORTS AND ENTERTAINMENT', 'HEALTH_CLUB_WEIGHT', 'Health club/weight rooms', 20.0, 0.06, 10.0, 2, ['health club', 'weight room'], 70),
      makeVentOcc('SPORTS AND ENTERTAINMENT', 'SPECTATOR_AREAS', 'Spectator areas', 7.5, 0.06, 150.0, 1, ['spectator'], 80),
      makeVentOcc('SPORTS AND ENTERTAINMENT', 'STAGES_STUDIOS', 'Stages, studios', 10.0, 0.06, 70.0, 1, ['stage', 'studio'], 90),
      makeVentOcc('SPORTS AND ENTERTAINMENT', 'SWIMMING_POOL_DECK', 'Swimming (pool and deck)', 0.0, 0.48, 0.0, 2, ['pool', 'deck', 'swimming'], 100),

      makeVentOcc('RESIDENTIAL', 'COMMON_CORRIDORS', 'Common corridors', 0.0, 0.06, 0.0, 1, ['residential corridor'], 10)
    ];
    ventilationOccupancyCatalog.sort(function (a, b) {
      var ga = String(a.category_group || '');
      var gb = String(b.category_group || '');
      if (ga < gb) return -1;
      if (ga > gb) return 1;
      var sa = parseNonNegativeNumber(a.sort_order, 0);
      var sb = parseNonNegativeNumber(b.sort_order, 0);
      if (sa < sb) return -1;
      if (sa > sb) return 1;
      var la = String(a.occupancy_category || '');
      var lb = String(b.occupancy_category || '');
      if (la < lb) return -1;
      if (la > lb) return 1;
      return 0;
    });
    var ventilationEzCatalog = [
      { key: 'DEFAULT_DESIGN_BASIS_08', label: 'Default design basis', ez: 0.8 },
      { key: 'COOLING_CEILING_SUPPLY', label: 'Ceiling supply of cool air', ez: 1.0 },
      { key: 'WARM_CEILING_FLOOR_RETURN', label: 'Ceiling supply of warm air and floor return', ez: 1.0 },
      { key: 'WARM_15PLUS_CEILING_RETURN', label: 'Ceiling supply warm air 15°F+ above space and ceiling return', ez: 0.8 },
      { key: 'WARM_LOW_THROW', label: 'Ceiling supply warm air <15°F above space, low throw, ceiling return', ez: 0.8 },
      { key: 'WARM_HIGH_THROW', label: 'Ceiling supply warm air <15°F above space, higher throw, ceiling return', ez: 1.0 },
      { key: 'FLOOR_WARM_FLOOR_RETURN', label: 'Floor supply of warm air and floor return', ez: 1.0 },
      { key: 'FLOOR_WARM_CEILING_RETURN', label: 'Floor supply of warm air and ceiling return', ez: 0.7 },
      { key: 'MAKEUP_GT_HALF', label: 'Makeup outlet > half room length from exhaust/return', ez: 0.8 },
      { key: 'MAKEUP_LT_HALF', label: 'Makeup outlet < half room length from exhaust/return', ez: 0.5 },
      { key: 'CUSTOM', label: 'Manual custom Ez', ez: 1.0 }
    ];
    var sanitary7032Table = {
      HORIZONTAL: [
        { size: '1-1/4"', maxUnits: 1, maxLength: 'Unlimited' },
        { size: '1-1/2"', maxUnits: 1, maxLength: 'Unlimited' },
        { size: '2"', maxUnits: 8, maxLength: 'Unlimited' },
        { size: '3"', maxUnits: 35, maxLength: 'Unlimited' },
        { size: '4"', maxUnits: 216, maxLength: 'Unlimited' },
        { size: '5"', maxUnits: 428, maxLength: 'Unlimited' },
        { size: '6"', maxUnits: 720, maxLength: 'Unlimited' },
        { size: '8"', maxUnits: 2640, maxLength: 'Unlimited' },
        { size: '10"', maxUnits: 4680, maxLength: 'Unlimited' },
        { size: '12"', maxUnits: 8200, maxLength: 'Unlimited' }
      ],
      VERTICAL: [
        { size: '1-1/4"', maxUnits: 1, maxLength: 45 },
        { size: '1-1/2"', maxUnits: 2, maxLength: 65 },
        { size: '2"', maxUnits: 16, maxLength: 85 },
        { size: '3"', maxUnits: 48, maxLength: 212 },
        { size: '4"', maxUnits: 256, maxLength: 300 },
        { size: '5"', maxUnits: 600, maxLength: 390 },
        { size: '6"', maxUnits: 1380, maxLength: 510 },
        { size: '8"', maxUnits: 3600, maxLength: 750 },
        { size: '10"', maxUnits: 5600, maxLength: '' },
        { size: '12"', maxUnits: 8400, maxLength: '' }
      ]
    };
    var gasPipeData = {
      IPC: {
        NATURAL_GAS: {
          LOW: [
            { length: 20, capacities: { '1/2"': 175, '3/4"': 360, '1"': 680, '1-1/4"': 1400, '1-1/2"': 2100, '2"': 3950, '2-1/2"': 6300, '3"': 11000 } },
            { length: 40, capacities: { '1/2"': 120, '3/4"': 250, '1"': 465, '1-1/4"': 950, '1-1/2"': 1460, '2"': 2750, '2-1/2"': 4350, '3"': 7700 } },
            { length: 60, capacities: { '1/2"': 97, '3/4"': 200, '1"': 375, '1-1/4"': 770, '1-1/2"': 1180, '2"': 2220, '2-1/2"': 3520, '3"': 6250 } },
            { length: 100, capacities: { '1/2"': 74, '3/4"': 152, '1"': 285, '1-1/4"': 590, '1-1/2"': 900, '2"': 1680, '2-1/2"': 2680, '3"': 4760 } }
          ],
          MEDIUM: [
            { length: 20, capacities: { '1/2"': 1200, '3/4"': 2470, '1"': 4660, '1-1/4"': 9600, '1-1/2"': 14500, '2"': 27300, '2-1/2"': 43300, '3"': 77000 } },
            { length: 40, capacities: { '1/2"': 1020, '3/4"': 2090, '1"': 3940, '1-1/4"': 8100, '1-1/2"': 12300, '2"': 23100, '2-1/2"': 36600, '3"': 65000 } },
            { length: 60, capacities: { '1/2"': 930, '3/4"': 1900, '1"': 3590, '1-1/4"': 7400, '1-1/2"': 11200, '2"': 21000, '2-1/2"': 33300, '3"': 59200 } },
            { length: 100, capacities: { '1/2"': 780, '3/4"': 1610, '1"': 3040, '1-1/4"': 6250, '1-1/2"': 9500, '2"': 17800, '2-1/2"': 28200, '3"': 50100 } }
          ],
          HIGH: [
            { length: 20, capacities: { '1/2"': 1900, '3/4"': 3920, '1"': 7390, '1-1/4"': 15200, '1-1/2"': 23000, '2"': 43300, '2-1/2"': 68700, '3"': 122000 } },
            { length: 40, capacities: { '1/2"': 1700, '3/4"': 3520, '1"': 6630, '1-1/4"': 13600, '1-1/2"': 20600, '2"': 38800, '2-1/2"': 61500, '3"': 109000 } },
            { length: 60, capacities: { '1/2"': 1560, '3/4"': 3230, '1"': 6090, '1-1/4"': 12500, '1-1/2"': 18900, '2"': 35600, '2-1/2"': 56500, '3"': 100000 } },
            { length: 100, capacities: { '1/2"': 1360, '3/4"': 2810, '1"': 5290, '1-1/4"': 10900, '1-1/2"': 16500, '2"': 31000, '2-1/2"': 49200, '3"': 87300 } }
          ]
        },
        PROPANE: {
          LOW: [
            { length: 20, capacities: { '1/2"': 390, '3/4"': 810, '1"': 1520, '1-1/4"': 3140, '1-1/2"': 4720, '2"': 8900, '2-1/2"': 14100, '3"': 25000 } },
            { length: 40, capacities: { '1/2"': 270, '3/4"': 560, '1"': 1060, '1-1/4"': 2180, '1-1/2"': 3320, '2"': 6200, '2-1/2"': 9800, '3"': 17400 } },
            { length: 60, capacities: { '1/2"': 220, '3/4"': 460, '1"': 870, '1-1/4"': 1790, '1-1/2"': 2720, '2"': 5080, '2-1/2"': 8050, '3"': 14300 } },
            { length: 100, capacities: { '1/2"': 170, '3/4"': 350, '1"': 670, '1-1/4"': 1370, '1-1/2"': 2100, '2"': 3910, '2-1/2"': 6190, '3"': 11000 } }
          ],
          MEDIUM: [
            { length: 20, capacities: { '1/2"': 2680, '3/4"': 5530, '1"': 10430, '1-1/4"': 21400, '1-1/2"': 32400, '2"': 60900, '2-1/2"': 96800, '3"': 172000 } },
            { length: 40, capacities: { '1/2"': 2280, '3/4"': 4700, '1"': 8870, '1-1/4"': 18200, '1-1/2"': 27600, '2"': 51800, '2-1/2"': 82300, '3"': 146000 } },
            { length: 60, capacities: { '1/2"': 2080, '3/4"': 4290, '1"': 8090, '1-1/4"': 16600, '1-1/2"': 25200, '2"': 47100, '2-1/2"': 74800, '3"': 133000 } },
            { length: 100, capacities: { '1/2"': 1740, '3/4"': 3590, '1"': 6770, '1-1/4"': 13900, '1-1/2"': 21100, '2"': 39400, '2-1/2"': 62500, '3"': 111000 } }
          ],
          HIGH: [
            { length: 20, capacities: { '1/2"': 4250, '3/4"': 8760, '1"': 16500, '1-1/4"': 34000, '1-1/2"': 51400, '2"': 96800, '2-1/2"': 153000, '3"': 273000 } },
            { length: 40, capacities: { '1/2"': 3800, '3/4"': 7840, '1"': 14750, '1-1/4"': 30300, '1-1/2"': 45900, '2"': 86400, '2-1/2"': 137000, '3"': 243000 } },
            { length: 60, capacities: { '1/2"': 3490, '3/4"': 7190, '1"': 13530, '1-1/4"': 27800, '1-1/2"': 42200, '2"': 79400, '2-1/2"': 126000, '3"': 223000 } },
            { length: 100, capacities: { '1/2"': 3040, '3/4"': 6260, '1"': 11770, '1-1/4"': 24200, '1-1/2"': 36700, '2"': 69000, '2-1/2"': 110000, '3"': 194000 } }
          ]
        }
      },
      CPC: {
        NATURAL_GAS: {
          LOW: [
            { length: 20, capacities: { '1/2"': 165, '3/4"': 340, '1"': 640, '1-1/4"': 1320, '1-1/2"': 1980, '2"': 3720, '2-1/2"': 5920, '3"': 10400 } },
            { length: 40, capacities: { '1/2"': 112, '3/4"': 232, '1"': 430, '1-1/4"': 880, '1-1/2"': 1340, '2"': 2520, '2-1/2"': 4020, '3"': 7060 } },
            { length: 60, capacities: { '1/2"': 90, '3/4"': 186, '1"': 345, '1-1/4"': 705, '1-1/2"': 1080, '2"': 2030, '2-1/2"': 3240, '3"': 5690 } },
            { length: 100, capacities: { '1/2"': 68, '3/4"': 140, '1"': 260, '1-1/4"': 535, '1-1/2"': 820, '2"': 1540, '2-1/2"': 2460, '3"': 4320 } }
          ],
          MEDIUM: [
            { length: 20, capacities: { '1/2"': 1140, '3/4"': 2350, '1"': 4420, '1-1/4"': 9090, '1-1/2"': 13800, '2"': 25800, '2-1/2"': 40900, '3"': 72700 } },
            { length: 40, capacities: { '1/2"': 970, '3/4"': 1990, '1"': 3740, '1-1/4"': 7690, '1-1/2"': 11700, '2"': 21900, '2-1/2"': 34800, '3"': 61800 } },
            { length: 60, capacities: { '1/2"': 885, '3/4"': 1810, '1"': 3410, '1-1/4"': 7000, '1-1/2"': 10600, '2"': 19800, '2-1/2"': 31500, '3"': 55900 } },
            { length: 100, capacities: { '1/2"': 740, '3/4"': 1540, '1"': 2890, '1-1/4"': 5940, '1-1/2"': 9000, '2"': 16900, '2-1/2"': 26800, '3"': 47500 } }
          ],
          HIGH: [
            { length: 20, capacities: { '1/2"': 1810, '3/4"': 3730, '1"': 7040, '1-1/4"': 14500, '1-1/2"': 21900, '2"': 41000, '2-1/2"': 65300, '3"': 116000 } },
            { length: 40, capacities: { '1/2"': 1620, '3/4"': 3350, '1"': 6320, '1-1/4"': 13000, '1-1/2"': 19600, '2"': 36700, '2-1/2"': 58300, '3"': 103000 } },
            { length: 60, capacities: { '1/2"': 1490, '3/4"': 3070, '1"': 5790, '1-1/4"': 11900, '1-1/2"': 18000, '2"': 33600, '2-1/2"': 53400, '3"': 94400 } },
            { length: 100, capacities: { '1/2"': 1290, '3/4"': 2670, '1"': 5030, '1-1/4"': 10400, '1-1/2"': 15700, '2"': 29400, '2-1/2"': 46700, '3"': 82600 } }
          ]
        },
        PROPANE: {
          LOW: [
            { length: 20, capacities: { '1/2"': 370, '3/4"': 770, '1"': 1450, '1-1/4"': 2990, '1-1/2"': 4490, '2"': 8450, '2-1/2"': 13400, '3"': 23800 } },
            { length: 40, capacities: { '1/2"': 255, '3/4"': 535, '1"': 1010, '1-1/4"': 2080, '1-1/2"': 3170, '2"': 5920, '2-1/2"': 9390, '3"': 16600 } },
            { length: 60, capacities: { '1/2"': 210, '3/4"': 440, '1"': 830, '1-1/4"': 1710, '1-1/2"': 2600, '2"': 4860, '2-1/2"': 7710, '3"': 13700 } },
            { length: 100, capacities: { '1/2"': 162, '3/4"': 338, '1"': 638, '1-1/4"': 1310, '1-1/2"': 2010, '2"': 3760, '2-1/2"': 5960, '3"': 10600 } }
          ],
          MEDIUM: [
            { length: 20, capacities: { '1/2"': 2550, '3/4"': 5260, '1"': 9920, '1-1/4"': 20400, '1-1/2"': 30900, '2"': 58100, '2-1/2"': 92300, '3"': 164000 } },
            { length: 40, capacities: { '1/2"': 2170, '3/4"': 4470, '1"': 8420, '1-1/4"': 17300, '1-1/2"': 26300, '2"': 49300, '2-1/2"': 78400, '3"': 139000 } },
            { length: 60, capacities: { '1/2"': 1980, '3/4"': 4080, '1"': 7680, '1-1/4"': 15800, '1-1/2"': 24000, '2"': 44900, '2-1/2"': 71300, '3"': 127000 } },
            { length: 100, capacities: { '1/2"': 1660, '3/4"': 3420, '1"': 6430, '1-1/4"': 13200, '1-1/2"': 20100, '2"': 37600, '2-1/2"': 59800, '3"': 106000 } }
          ],
          HIGH: [
            { length: 20, capacities: { '1/2"': 4040, '3/4"': 8340, '1"': 15700, '1-1/4"': 32400, '1-1/2"': 49000, '2"': 92300, '2-1/2"': 146000, '3"': 260000 } },
            { length: 40, capacities: { '1/2"': 3610, '3/4"': 7470, '1"': 14060, '1-1/4"': 28900, '1-1/2"': 43800, '2"': 82300, '2-1/2"': 130000, '3"': 232000 } },
            { length: 60, capacities: { '1/2"': 3310, '3/4"': 6850, '1"': 12910, '1-1/4"': 26600, '1-1/2"': 40300, '2"': 75700, '2-1/2"': 120000, '3"': 213000 } },
            { length: 100, capacities: { '1/2"': 2890, '3/4"': 5950, '1"': 11220, '1-1/4"': 23100, '1-1/2"': 35100, '2"': 65900, '2-1/2"': 104000, '3"': 186000 } }
          ]
        }
      }
    };
    var gasTableCatalog = [];
    var gasDrafts = [];
    var GAS_PIPE_SIZES = ['1/2"', '3/4"', '1"', '1-1/4"', '1-1/2"', '2"', '2-1/2"', '3"'];
    var GAS_MATERIAL_OPTIONS = [
      { id: 'SCHEDULE_40_METALLIC', label: 'Schedule 40 Metallic Pipe', factor: 1.00 },
      { id: 'CSST', label: 'CSST', factor: 0.92 },
      { id: 'COPPER', label: 'Copper Tubing', factor: 0.88 },
      { id: 'PE', label: 'PE (Polyethylene)', factor: 0.95 },
      { id: 'GALVANIZED_STEEL', label: 'Galvanized Steel', factor: 0.98 },
      { id: 'WROUGHT_IRON', label: 'Wrought Iron', factor: 1.00 }
    ];

    function getDefaultSpecificGravityForFuel(fuelType) {
      return fuelType === 'PROPANE' ? 1.52 : 0.60;
    }

    function getGasMaterialLabel(materialId) {
      var i;
      for (i = 0; i < GAS_MATERIAL_OPTIONS.length; i++) {
        if (GAS_MATERIAL_OPTIONS[i].id === materialId) return GAS_MATERIAL_OPTIONS[i].label;
      }
      return 'Schedule 40 Metallic Pipe';
    }

    function getGasMaterialFactor(materialId) {
      var i;
      for (i = 0; i < GAS_MATERIAL_OPTIONS.length; i++) {
        if (GAS_MATERIAL_OPTIONS[i].id === materialId) return GAS_MATERIAL_OPTIONS[i].factor;
      }
      return 1.0;
    }

    function buildRowsFromLegacyTable(entries) {
      var rows = [];
      var i, key;
      for (i = 0; i < entries.length; i++) {
        var caps = {};
        for (key in entries[i].capacities) {
          caps[key] = parseNonNegativeNumber(entries[i].capacities[key], 0);
        }
        rows.push({
          length: parseNonNegativeNumber(entries[i].length, 0),
          capacities: caps
        });
      }
      rows.sort(function (a, b) { return a.length - b.length; });
      return rows;
    }

    function rowsToLengthAndMatrix(rows) {
      var lengthRows = [];
      var sizeCapacityMatrix = {};
      var i, j, size;
      for (i = 0; i < GAS_PIPE_SIZES.length; i++) sizeCapacityMatrix[GAS_PIPE_SIZES[i]] = [];
      for (i = 0; i < rows.length; i++) {
        lengthRows.push(parseNonNegativeNumber(rows[i].length, 0));
        for (j = 0; j < GAS_PIPE_SIZES.length; j++) {
          size = GAS_PIPE_SIZES[j];
          sizeCapacityMatrix[size].push(parseNonNegativeNumber(rows[i].capacities[size], 0));
        }
      }
      return { lengthRows: lengthRows, sizeCapacityMatrix: sizeCapacityMatrix };
    }

    function matrixToRows(table) {
      var rows = [];
      var lengthRows = table.lengthRows || [];
      var matrix = table.sizeCapacityMatrix || {};
      var i, j, size, caps;
      for (i = 0; i < lengthRows.length; i++) {
        caps = {};
        for (j = 0; j < GAS_PIPE_SIZES.length; j++) {
          size = GAS_PIPE_SIZES[j];
          caps[size] = parseNonNegativeNumber((matrix[size] && matrix[size][i] !== undefined ? matrix[size][i] : 0), 0);
        }
        rows.push({ length: parseNonNegativeNumber(lengthRows[i], 0), capacities: caps });
      }
      rows.sort(function (a, b) { return a.length - b.length; });
      return rows;
    }

    function seedLegacyGasTables() {
      if (gasTableCatalog.length > 0) return;
      var codeBasis, fuelType, pressureGroup, entries;
      var counter = 1;
      var pressureInfo = {
        LOW: { inletPressure: 7, pressureDrop: 0.3 },
        MEDIUM: { inletPressure: 2.0, pressureDrop: 1.0 },
        HIGH: { inletPressure: 5.0, pressureDrop: 2.0 }
      };
      for (codeBasis in gasPipeData) {
        for (fuelType in gasPipeData[codeBasis]) {
          for (pressureGroup in gasPipeData[codeBasis][fuelType]) {
            entries = gasPipeData[codeBasis][fuelType][pressureGroup];
            var rowResult = rowsToLengthAndMatrix(buildRowsFromLegacyTable(entries));
            var matIndex, materialId, factor, scaledMatrix, sizeKey, capIndex;
            for (matIndex = 0; matIndex < GAS_MATERIAL_OPTIONS.length; matIndex++) {
              materialId = GAS_MATERIAL_OPTIONS[matIndex].id;
              factor = getGasMaterialFactor(materialId);
              scaledMatrix = {};
              for (sizeKey in rowResult.sizeCapacityMatrix) {
                scaledMatrix[sizeKey] = [];
                for (capIndex = 0; capIndex < rowResult.sizeCapacityMatrix[sizeKey].length; capIndex++) {
                  scaledMatrix[sizeKey].push(roundNumber(parseNonNegativeNumber(rowResult.sizeCapacityMatrix[sizeKey][capIndex], 0) * factor, 0));
                }
              }
              gasTableCatalog.push({
                id: 'seed_' + counter,
                gasType: fuelType,
                inletPressure: pressureInfo[pressureGroup].inletPressure,
                pressureDrop: pressureInfo[pressureGroup].pressureDrop,
                specificGravity: getDefaultSpecificGravityForFuel(fuelType),
                material: materialId,
                codeBasis: codeBasis,
                sourceFile: 'Built-in starter dataset',
                tableName: codeBasis + ' ' + (fuelType === 'NATURAL_GAS' ? 'Natural Gas' : 'Propane') + ' ' + pressureGroup + ' - ' + getGasMaterialLabel(materialId),
                lengthRows: rowResult.lengthRows,
                sizeCapacityMatrix: scaledMatrix,
                rawPreview: 'Seeded from in-app starter data.'
              });
              counter += 1;
            }
          }
        }
      }
    }

    var DUCT_CALC = {
      modes: {
        SIZE_FROM_CFM_FRICTION: 'SIZE_FROM_CFM_FRICTION',
        SIZE_FROM_CFM_VELOCITY: 'SIZE_FROM_CFM_VELOCITY',
        CHECK_ROUND_SIZE: 'CHECK_ROUND_SIZE',
        CHECK_RECT_SIZE: 'CHECK_RECT_SIZE'
      },
      shapes: { ROUND: 'ROUND', RECTANGULAR: 'RECTANGULAR' },
      standardRoundIn: [4,5,6,7,8,9,10,11,12,13,14,15,16,18,20,22,24,26,28,30,32,34,36,38,40,42,44,46,48,50,54,60],
      standardRectIn: [4,5,6,7,8,9,10,11,12,13,14,15,16,18,20,22,24,26,28,30,32,34,36,38,40,42,44,46,48,50,54,60],
      calibrationTolerancePct: 2.0,
      calibrationCases: [] // Populate with DuctSizer sample I/O pairs for exact parity tuning.
    };
    var DUCT_STATIC_C_TABLE = {
      rwBins: [0.5, 1.0, 1.5, 2.0, 2.5, 3.0, 4.0],
      hwBins: [0.5, 1.0, 1.5, 2.0, 3.0, 4.0, 6.0],
      // Approximate CR3-1 elbow-loss matrix for baseline; tune with your approved table values.
      values: [
        [1.45, 1.35, 1.25, 1.15, 1.05, 0.98, 0.92],
        [1.30, 1.20, 1.12, 1.04, 0.96, 0.90, 0.86],
        [1.18, 1.10, 1.02, 0.95, 0.89, 0.84, 0.80],
        [1.08, 1.00, 0.93, 0.87, 0.82, 0.78, 0.75],
        [0.99, 0.92, 0.86, 0.81, 0.77, 0.73, 0.70],
        [0.92, 0.86, 0.81, 0.76, 0.72, 0.69, 0.66],
        [0.86, 0.80, 0.76, 0.71, 0.68, 0.65, 0.62]
      ]
    };

    function ductAreaSqFt(shape, diameterIn, widthIn, heightIn) {
      if (shape === DUCT_CALC.shapes.ROUND) {
        return (Math.PI * Math.pow(diameterIn, 2) / 4.0) / 144.0;
      }
      return (widthIn * heightIn) / 144.0;
    }

    function ductEquivalentDiameterIn(shape, diameterIn, widthIn, heightIn) {
      if (shape === DUCT_CALC.shapes.ROUND) return diameterIn;
      if (widthIn <= 0 || heightIn <= 0) return 0;
      return 1.3 * Math.pow(widthIn * heightIn, 0.625) / Math.pow(widthIn + heightIn, 0.25);
    }

    function ductVelocityFpm(cfm, shape, diameterIn, widthIn, heightIn) {
      var area = ductAreaSqFt(shape, diameterIn, widthIn, heightIn);
      if (area <= 0) return 0;
      return cfm / area;
    }

    function ductVelocityPressure(velocityFpm) {
      return Math.pow(velocityFpm / 4005.0, 2);
    }

    // Baseline HVAC duct friction model (in.wg per 100ft), to be tuned with DuctSizer sample calibration.
    function ductFrictionRate(cfm, eqDiameterIn) {
      if (cfm <= 0 || eqDiameterIn <= 0) return 0;
      return 0.109136 * Math.pow(cfm, 1.9) / Math.pow(eqDiameterIn, 5.02);
    }

    function ductRoundSizeFromTarget(cfm, targetFriction) {
      if (cfm <= 0 || targetFriction <= 0) return 0;
      var d = Math.pow((0.109136 * Math.pow(cfm, 1.9)) / targetFriction, 1.0 / 5.02);
      return d;
    }

    function ductChooseStandardRound(minDiameter) {
      var i;
      for (i = 0; i < DUCT_CALC.standardRoundIn.length; i++) {
        if (DUCT_CALC.standardRoundIn[i] >= minDiameter) return DUCT_CALC.standardRoundIn[i];
      }
      return DUCT_CALC.standardRoundIn[DUCT_CALC.standardRoundIn.length - 1];
    }

    function ductChooseRectFromEquivalent(targetEqDiameter, ratioLimit) {
      var widths = DUCT_CALC.standardRectIn;
      var heights = DUCT_CALC.standardRectIn;
      var best = null;
      var wi, hi, w, h, ratio, eq, area;
      if (targetEqDiameter <= 0) return null;
      if (!ratioLimit || ratioLimit < 1) ratioLimit = 3.0;
      for (wi = 0; wi < widths.length; wi++) {
        for (hi = 0; hi < heights.length; hi++) {
          w = widths[wi];
          h = heights[hi];
          if (w < h) continue;
          ratio = w / h;
          if (ratio > ratioLimit) continue;
          eq = ductEquivalentDiameterIn(DUCT_CALC.shapes.RECTANGULAR, 0, w, h);
          if (eq < targetEqDiameter) continue;
          area = w * h;
          if (!best || area < best.area || (area === best.area && ratio < best.ratio)) {
            best = { width: w, height: h, ratio: ratio, equivalentDiameter: eq, area: area };
          }
        }
      }
      return best;
    }

    function normalizeDuctMode(mode, shapeHint) {
      var shape = shapeHint || DUCT_CALC.shapes.ROUND;
      if (mode === DUCT_CALC.modes.SIZE_FROM_CFM_FRICTION ||
          mode === DUCT_CALC.modes.SIZE_FROM_CFM_VELOCITY ||
          mode === DUCT_CALC.modes.CHECK_ROUND_SIZE ||
          mode === DUCT_CALC.modes.CHECK_RECT_SIZE) {
        return mode;
      }
      if (mode === 'CHECK_VELOCITY_PRESSURE' || mode === 'CHECK_FRICTION') {
        return (shape === DUCT_CALC.shapes.RECTANGULAR) ? DUCT_CALC.modes.CHECK_RECT_SIZE : DUCT_CALC.modes.CHECK_ROUND_SIZE;
      }
      return DUCT_CALC.modes.SIZE_FROM_CFM_FRICTION;
    }

    function getDuctModeLabel(mode) {
      if (mode === DUCT_CALC.modes.SIZE_FROM_CFM_VELOCITY) return 'Size from CFM + Velocity';
      if (mode === DUCT_CALC.modes.CHECK_ROUND_SIZE) return 'Check from Round Size + CFM';
      if (mode === DUCT_CALC.modes.CHECK_RECT_SIZE) return 'Check from Rectangular Size + CFM';
      return 'Size from CFM + Friction';
    }

    function getDuctSelectionState() {
      var currentShape = document.getElementById('ductShape') ? document.getElementById('ductShape').value : DUCT_CALC.shapes.ROUND;
      var currentMode = document.getElementById('ductMode') ? document.getElementById('ductMode').value : DUCT_CALC.modes.SIZE_FROM_CFM_FRICTION;
      currentMode = normalizeDuctMode(currentMode, currentShape);
      return {
        mode: currentMode,
        shape: currentShape,
        cfm: parseNonNegativeNumber(document.getElementById('ductCfm') ? document.getElementById('ductCfm').value : 0, 0),
        frictionTarget: parseNonNegativeNumber(document.getElementById('ductFrictionTarget') ? document.getElementById('ductFrictionTarget').value : 0, 0),
        velocityTarget: parseNonNegativeNumber(document.getElementById('ductVelocityTarget') ? document.getElementById('ductVelocityTarget').value : 0, 0),
        diameter: parseNonNegativeNumber(document.getElementById('ductDiameter') ? document.getElementById('ductDiameter').value : 0, 0),
        width: parseNonNegativeNumber(document.getElementById('ductWidth') ? document.getElementById('ductWidth').value : 0, 0),
        height: parseNonNegativeNumber(document.getElementById('ductHeight') ? document.getElementById('ductHeight').value : 0, 0),
        ratioLimit: parseNonNegativeNumber(document.getElementById('ductRatioLimit') ? document.getElementById('ductRatioLimit').value : 3, 3)
      };
    }

    function syncDuctSegmentedControls() {
      var mode = document.getElementById('ductMode') ? document.getElementById('ductMode').value : DUCT_CALC.modes.SIZE_FROM_CFM_FRICTION;
      var shape = document.getElementById('ductShape') ? document.getElementById('ductShape').value : DUCT_CALC.shapes.ROUND;
      var modeButtons = document.querySelectorAll('[data-duct-mode]');
      var shapeButtons = document.querySelectorAll('[data-duct-shape]');
      var i;
      for (i = 0; i < modeButtons.length; i++) {
        modeButtons[i].classList.toggle('active', modeButtons[i].getAttribute('data-duct-mode') === mode);
      }
      for (i = 0; i < shapeButtons.length; i++) {
        shapeButtons[i].classList.toggle('active', shapeButtons[i].getAttribute('data-duct-shape') === shape);
      }
    }

    function setDuctMode(mode) {
      var modeNode = document.getElementById('ductMode');
      var shapeNode = document.getElementById('ductShape');
      if (!modeNode) return;
      modeNode.value = normalizeDuctMode(mode, shapeNode ? shapeNode.value : DUCT_CALC.shapes.ROUND);
      if (shapeNode) {
        if (modeNode.value === DUCT_CALC.modes.CHECK_ROUND_SIZE) shapeNode.value = DUCT_CALC.shapes.ROUND;
        if (modeNode.value === DUCT_CALC.modes.CHECK_RECT_SIZE) shapeNode.value = DUCT_CALC.shapes.RECTANGULAR;
      }
      onDuctInputsChanged();
    }

    function setDuctShape(shape) {
      var shapeNode = document.getElementById('ductShape');
      var modeNode = document.getElementById('ductMode');
      if (!shapeNode || !modeNode) return;
      if (modeNode.value !== DUCT_CALC.modes.SIZE_FROM_CFM_FRICTION && modeNode.value !== DUCT_CALC.modes.SIZE_FROM_CFM_VELOCITY) return;
      shapeNode.value = shape;
      onDuctInputsChanged();
    }

    function updateHomeDashboardFit() {
      var body = document.body;
      var isHome = (activeModule === 'home');
      var viewportH;
      var wrap;
      var requiresTight;
      var requiresExtraTight;
      if (!body) return;
      if (!isHome) {
        body.classList.remove('home-dashboard-fit');
        body.classList.remove('home-tight');
        body.classList.remove('home-tight-plus');
        return;
      }
      body.classList.add('home-dashboard-fit');
      viewportH = window.innerHeight || document.documentElement.clientHeight || 0;
      requiresTight = (viewportH > 0 && viewportH < 860);
      wrap = document.getElementsByClassName('wrap')[0] || null;
      if (!requiresTight && wrap) {
        requiresTight = (wrap.scrollHeight > (wrap.clientHeight + 1));
      }
      if (requiresTight) body.classList.add('home-tight');
      else body.classList.remove('home-tight');
      requiresExtraTight = false;
      if (wrap) {
        requiresExtraTight = (wrap.scrollHeight > (wrap.clientHeight + 1));
      }
      if (requiresExtraTight) body.classList.add('home-tight-plus');
      else body.classList.remove('home-tight-plus');
    }

    function setActiveModule(moduleName) {
      var screens = document.getElementsByClassName('module-screen');
      var navButtons = document.getElementsByClassName('nav-btn');
      var i;
      activeModule = moduleName;

      for (i = 0; i < screens.length; i++) {
        screens[i].className = screens[i].className.replace(/\bactive\b/g, '').replace(/\s{2,}/g, ' ').trim();
      }
      for (i = 0; i < navButtons.length; i++) {
        navButtons[i].className = navButtons[i].className.replace(/\bactive\b/g, '').replace(/\s{2,}/g, ' ').trim();
      }

      var screen = document.getElementById(moduleName + 'Screen');
      var nav = document.getElementById('nav' + moduleName.charAt(0).toUpperCase() + moduleName.slice(1));
      if (screen) screen.className += ' active';
      if (nav) nav.className += ' active';
      if (moduleName === 'home') {
        renderCurrentProjectCard();
        syncAppMetaFromBackend(0);
      }
      updateHomeDashboardFit();
    }

    function openModule(moduleName) {
      setActiveModule(moduleName);
      ensureModuleRendered(moduleName);
      if (moduleName === 'storm') recalculate(false);
    }

    function ensureModuleRendered(moduleName) {
      try {
        if (moduleName === 'condensate') {
          if (!condensateZones || condensateZones.length <= 0) resetCondensateSection();
          else { renderCondensateZones(); updateCondensatePreview(); }
        } else if (moduleName === 'vent') {
          if (!sanitaryZones || sanitaryZones.length <= 0) resetVentSection();
          else updateVentTotalsAndUI();
        } else if (moduleName === 'wsfu') {
          if (!wsfuZones || wsfuZones.length <= 0) resetWsfuSection();
          else updateWsfuTotalsAndUI();
        } else if (moduleName === 'fixtureUnit') {
          if (!fixtureUnitZones || fixtureUnitZones.length <= 0) fixtureUnitZones = [createDefaultFixtureUnitZone(1)];
          renderFixtureUnitRows();
          updateFixtureUnitRowsAndPreview();
        } else if (moduleName === 'duct') {
          if (!ductZones || ductZones.length <= 0) resetDuctSection();
          else { renderDuctZones(); calculateDuct(false); }
        } else if (moduleName === 'ductStatic') {
          if (!ductStaticSegments || ductStaticSegments.length <= 0) resetDuctStaticSection();
          else renderDuctStaticSegments();
        } else if (moduleName === 'ventilation') {
          ensureVentilationInitialized();
          renderVentilationEzOptions();
          bindVentilationEntryLiveEvents();
          renderVentilationZonesTable();
          calculateVentilation(false);
        } else if (moduleName === 'refrigerant') {
          if (!refrigerantMultiRows || refrigerantMultiRows.length <= 0) addRefrigerantMultiRow();
          else renderRefrigerantMultiRows();
        } else if (moduleName === 'solar') {
          onSolarSaraToggle();
          updateSolarPreview();
        } else if (moduleName === 'gas') {
          if (!document.getElementById('gasLinesBody') || document.getElementById('gasLinesBody').innerHTML.trim() === '') {
            clearGasLineRows();
            addGasLineRow({ id: makeGasLineId(), label: 'Line 1', demandValue: 0, runLength: 0 });
          }
        }
      } catch (e) {
        console.error('Module render restore failed for ' + moduleName, e);
      }
    }

    function backToHome() {
      setActiveModule('home');
    }

    function roundNumber(value, digits) {
      var factor = Math.pow(10, digits);
      return Math.round(value * factor) / factor;
    }

    function formatNumber(value, digits) {
      var n = roundNumber(value, digits);
      var parts = n.toFixed(digits).split('.');
      parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ",");
      return parts.join('.');
    }

    function getAreaBlocks() {
      var container = document.getElementById('areasContainer');
      var divs, blocks, i;
      if (!container) return [];

      divs = container.getElementsByTagName('div');
      blocks = [];
      for (i = 0; i < divs.length; i++) {
        if ((' ' + divs[i].className + ' ').indexOf(' area-block ') >= 0) {
          blocks.push(divs[i]);
        }
      }
      return blocks;
    }

    function getAreaLabel(index) {
      if (index < 26) return 'Area ' + String.fromCharCode(65 + index);
      return 'Area ' + (index + 1);
    }

    function makeAreaBlock(areaLabel, defaultArea) {
      var html = '';
      html += '<div class="area-block">';
      html += '  <div class="area-header-row">';
      html += '    <div class="area-title-inline"><input type="text" class="area-name-input" value="' + areaLabel + '" oninput="onAreaNameChanged(this);" onchange="onAreaNameChanged(this);" /></div>';
      html += '    <div class="area-main-wrap"><input type="number" class="area-main-input" placeholder="Area in Square Feet" value="' + defaultArea + '" oninput="recalculate();" onchange="recalculate();" /></div>';
      html += '    <div class="area-actions">';
      html += '      <div class="area-meta-group">';
      html += '        <div class="subarea-count count-chip">Sub Areas: 0</div>';
      html += '        <div class="sidewall-count count-chip">Side Walls: 0</div>';
      html += '      </div>';
      html += '      <input type="button" class="btn btn-subarea btn-area-action storm-utility-btn" value="+ Sub Area" onclick="addSubArea(this);" />';
      html += '      <input type="button" class="btn btn-sidewall btn-area-action storm-utility-btn" value="+ Side Wall" onclick="addSideWall(this);" />';
      html += '      <input type="button" class="btn btn-small btn-area-action storm-utility-btn" value="&#128465; Remove" onclick="removeArea(this);" />';
      html += '    </div>';
      html += '  </div>';
      html += '  <div class="field-error area-main-error"></div>';
      html += '  <div class="subareas-container"></div>';
      html += '  <div class="sidewalls-container"></div>';
      html += '  <div class="field-error area-wall-error"></div>';
      html += '</div>';
      return html;
    }

    function makeSubAreaRow() {
      var html = '';
      html += '<div class="subarea-row">';
      html += '  <div class="subarea-label">Sub Area</div>';
      html += '  <div class="subarea-input-wrap"><input type="number" class="subarea-input" placeholder="Area in Square Feet" value="0" oninput="recalculate();" onchange="recalculate();" /></div>';
      html += '  <div class="subarea-actions"><input type="button" class="btn btn-subremove btn-area-action" value="&#128465; Remove Sub" onclick="removeSubArea(this);" /></div>';
      html += '  <div class="field-error"></div>';
      html += '</div>';
      return html;
    }

    function makeSideWallRow() {
      var html = '';
      html += '<div class="sidewall-row">';
      html += '  <div class="sidewall-label">Side Wall</div>';
      html += '  <div class="sidewall-input-wrap"><input type="number" class="sidewall-input" placeholder="Side Wall Area (sq ft)" value="0" oninput="recalculate();" onchange="recalculate();" /></div>';
      html += '  <div class="sidewall-actions"><input type="button" class="btn btn-sideremove btn-area-action" value="&#128465; Remove Wall" onclick="removeSideWall(this);" /></div>';
      html += '  <div class="field-error"></div>';
      html += '</div>';
      return html;
    }

    function getSubAreaRows(block) {
      var divs = block.getElementsByTagName('div');
      var rows = [];
      var i;
      for (i = 0; i < divs.length; i++) {
        if ((' ' + divs[i].className + ' ').indexOf(' subarea-row ') >= 0) rows.push(divs[i]);
      }
      return rows;
    }

    function getSideWallRows(block) {
      var divs = block.getElementsByTagName('div');
      var rows = [];
      var i;
      for (i = 0; i < divs.length; i++) {
        if ((' ' + divs[i].className + ' ').indexOf(' sidewall-row ') >= 0) rows.push(divs[i]);
      }
      return rows;
    }

    function updateSubAreaMetadata(block) {
      var rows = getSubAreaRows(block);
      var countNodes = block.getElementsByClassName('subarea-count');
      var parentLabel = getAreaNameFromBlock(block, -1);
      var i, labels;
      for (i = 0; i < rows.length; i++) {
        labels = rows[i].getElementsByClassName('subarea-label');
        if (labels.length > 0) labels[0].innerHTML = parentLabel + '-' + (i + 1);
      }
      if (countNodes.length > 0) countNodes[0].innerHTML = 'Sub Areas: ' + rows.length;
    }

    function updateSideWallMetadata(block) {
      var rows = getSideWallRows(block);
      var countNodes = block.getElementsByClassName('sidewall-count');
      var i, labels;
      for (i = 0; i < rows.length; i++) {
        labels = rows[i].getElementsByClassName('sidewall-label');
        if (labels.length > 0) labels[0].innerHTML = 'Side Wall ' + (i + 1);
      }
      if (countNodes.length > 0) countNodes[0].innerHTML = 'Side Walls: ' + rows.length;
    }

    function renumberAreas() {
      var blocks = getAreaBlocks();
      var i, inputs;
      for (i = 0; i < blocks.length; i++) {
        inputs = blocks[i].getElementsByClassName('area-name-input');
        if (inputs.length > 0 && (!inputs[0].value || !String(inputs[0].value).trim())) inputs[0].value = getAreaLabel(i);
        updateSubAreaMetadata(blocks[i]);
        updateSideWallMetadata(blocks[i]);
      }
    }

    function getAreaNameFromBlock(block, areaIndex) {
      var defaultLabel = (areaIndex >= 0) ? getAreaLabel(areaIndex) : 'Area';
      var inputs = block ? block.getElementsByClassName('area-name-input') : [];
      if (inputs.length <= 0) return defaultLabel;
      var custom = String(inputs[0].value || '').trim();
      return custom || defaultLabel;
    }

    function onAreaNameChanged(node) {
      var block = node;
      while (block && (' ' + block.className + ' ').indexOf(' area-block ') < 0) block = block.parentNode;
      if (!block) return;
      updateSubAreaMetadata(block);
      recalculate(false);
    }

    function addArea() {
      var count = getAreaBlocks().length;
      var container = document.getElementById('areasContainer');
      container.insertAdjacentHTML('beforeend', makeAreaBlock(getAreaLabel(count), 0));
      renumberAreas();
      recalculate();
    }

    function addSubArea(btn) {
      var block = btn;
      var container, containers;
      while (block && (' ' + block.className + ' ').indexOf(' area-block ') < 0) {
        block = block.parentNode;
      }
      if (!block) return;
      containers = block.getElementsByClassName('subareas-container');
      if (containers.length <= 0) return;
      container = containers[0];
      container.insertAdjacentHTML('beforeend', makeSubAreaRow());
      updateSubAreaMetadata(block);
      recalculate();
    }

    function addSideWall(btn) {
      var block = btn;
      var container, containers;
      while (block && (' ' + block.className + ' ').indexOf(' area-block ') < 0) {
        block = block.parentNode;
      }
      if (!block) return;
      containers = block.getElementsByClassName('sidewalls-container');
      if (containers.length <= 0) return;
      container = containers[0];
      container.insertAdjacentHTML('beforeend', makeSideWallRow());
      updateSideWallMetadata(block);
      recalculate();
    }

    function removeSubArea(btn) {
      var row = btn;
      var block = btn;
      while (row && (' ' + row.className + ' ').indexOf(' subarea-row ') < 0) {
        row = row.parentNode;
      }
      while (block && (' ' + block.className + ' ').indexOf(' area-block ') < 0) {
        block = block.parentNode;
      }
      if (row && row.parentNode) row.parentNode.removeChild(row);
      if (block) updateSubAreaMetadata(block);
      recalculate();
    }

    function removeSideWall(btn) {
      var row = btn;
      var block = btn;
      while (row && (' ' + row.className + ' ').indexOf(' sidewall-row ') < 0) {
        row = row.parentNode;
      }
      while (block && (' ' + block.className + ' ').indexOf(' area-block ') < 0) {
        block = block.parentNode;
      }
      if (row && row.parentNode) row.parentNode.removeChild(row);
      if (block) updateSideWallMetadata(block);
      recalculate();
    }

    function removeArea(btn) {
      var blocks = getAreaBlocks();
      if (blocks.length <= 1) {
        alert('At least one area is required.');
        return;
      }
      var block = btn;
      while (block && (' ' + block.className + ' ').indexOf(' area-block ') < 0) {
        block = block.parentNode;
      }
      if (block && block.parentNode) {
        block.parentNode.removeChild(block);
        renumberAreas();
        recalculate();
      }
    }

    function pickPipeSizeByArea(requiredArea, rainfallRate, entries) {
      var i;
      var factor = rainfallRate * 0.0104;
      var allowableArea;
      for (i = 0; i < entries.length; i++) {
        allowableArea = entries[i].gpm / factor;
        if (requiredArea <= allowableArea) return entries[i];
      }
      return entries[entries.length - 1];
    }

    function pickPipeSizeWithOverflow(requiredArea, rainfallRate, entries) {
      var selected = pickPipeSizeByArea(requiredArea, rainfallRate, entries);
      var maxEntry = entries[entries.length - 1];
      var maxAllowableArea = maxEntry.gpm / (rainfallRate * 0.0104);
      return {
        pipe: selected,
        exceeded: requiredArea > maxAllowableArea,
        maxCapacity: maxEntry.gpm,
        maxAllowableArea: maxAllowableArea
      };
    }

    function clearInputError(input, errorMessageNode) {
      input.className = input.className.replace(/\binput-error\b/g, '').replace(/\s{2,}/g, ' ');
      if (errorMessageNode) errorMessageNode.innerHTML = '';
    }

    function showInputError(input, errorMessageNode, message) {
      if (input.className.indexOf('input-error') < 0) input.className += ' input-error';
      if (errorMessageNode) errorMessageNode.innerHTML = message;
    }

    function getFieldErrorNode(block) {
      var divs = block.getElementsByTagName('div');
      var i;
      for (i = 0; i < divs.length; i++) {
        if (divs[i].className && divs[i].className.indexOf('field-error') >= 0) return divs[i];
      }
      return null;
    }

    function getSideWallFactor(sideWallCount) {
      if (sideWallCount == 1) return 0.50;
      if (sideWallCount == 2) return 0.35;
      if (sideWallCount >= 3) return 0.50;
      return 0.00;
    }

    function parseAreasForRunoff(rainfallRate) {
      var blocks = getAreaBlocks();
      var i, j, areaInput, areaErrorNode, raw, areaValue, areaLabel;
      var sideWallRows, sideWallInput, sideWallErrorNode, sideWallRaw, sideWallValue;
      var subRows, subInput, subErrorNode, subRaw, subValue;
      var totalArea = 0;
      var allValid = true;
      var perArea = [];
      var roofAreaTotal;
      var parentArea;
      var subAreaValues;
      var effectiveArea;
      var sideWallAreas;
      var sideWallSum;
      var sideWallFactor;
      var sideWallContribution;
      var wallErrorNode;

      for (i = 0; i < blocks.length; i++) {
        roofAreaTotal = 0;
        parentArea = 0;
        subAreaValues = [];
        sideWallAreas = [];
        sideWallSum = 0;
        sideWallFactor = 0;
        sideWallContribution = 0;
        areaInput = blocks[i].getElementsByClassName('area-main-input')[0];
        areaErrorNode = blocks[i].getElementsByClassName('area-main-error')[0];
        wallErrorNode = blocks[i].getElementsByClassName('area-wall-error')[0];
        if (!areaInput) continue;
        raw = areaInput.value;
        clearInputError(areaInput, areaErrorNode);
        if (wallErrorNode) wallErrorNode.innerHTML = '';

        if (raw === '' || isNaN(parseFloat(raw)) || parseFloat(raw) < 0) {
          allValid = false;
          showInputError(areaInput, areaErrorNode, 'Enter an area value of 0 or greater.');
        } else {
          areaValue = parseFloat(raw);
          parentArea = areaValue;
          roofAreaTotal += areaValue;
        }

        subRows = getSubAreaRows(blocks[i]);
        for (j = 0; j < subRows.length; j++) {
          subInput = subRows[j].getElementsByClassName('subarea-input')[0];
          subErrorNode = subRows[j].getElementsByClassName('field-error')[0];
          if (!subInput) continue;
          subRaw = subInput.value;
          clearInputError(subInput, subErrorNode);
          if (subRaw === '' || isNaN(parseFloat(subRaw)) || parseFloat(subRaw) < 0) {
            allValid = false;
            showInputError(subInput, subErrorNode, 'Enter a sub-area value of 0 or greater.');
          } else {
            subValue = parseFloat(subRaw);
            subAreaValues.push(subValue);
            roofAreaTotal += subValue;
          }
        }

        sideWallRows = getSideWallRows(blocks[i]);
        for (j = 0; j < sideWallRows.length; j++) {
          sideWallInput = sideWallRows[j].getElementsByClassName('sidewall-input')[0];
          sideWallErrorNode = sideWallRows[j].getElementsByClassName('field-error')[0];
          if (!sideWallInput) continue;
          sideWallRaw = sideWallInput.value;
          clearInputError(sideWallInput, sideWallErrorNode);
          if (sideWallRaw === '' || isNaN(parseFloat(sideWallRaw)) || parseFloat(sideWallRaw) < 0) {
            allValid = false;
            showInputError(sideWallInput, sideWallErrorNode, 'Enter a side wall area of 0 or greater.');
          } else {
            sideWallValue = parseFloat(sideWallRaw);
            sideWallAreas.push(sideWallValue);
            sideWallSum += sideWallValue;
          }
        }

        sideWallFactor = getSideWallFactor(sideWallAreas.length);
        sideWallContribution = sideWallSum * sideWallFactor;
        effectiveArea = roofAreaTotal + sideWallContribution;
        totalArea += effectiveArea;
        areaLabel = getAreaNameFromBlock(blocks[i], i);
        perArea.push({
          label: areaLabel,
          area: effectiveArea,
          gpm: rainfallRate * effectiveArea * 0.0104,
          roofArea: roofAreaTotal,
          sideWallAreas: sideWallAreas,
          sideWallCount: sideWallAreas.length,
          sideWallSum: sideWallSum,
          sideWallFactor: sideWallFactor,
          sideWallContribution: sideWallContribution,
          effectiveArea: effectiveArea,
          parentArea: parentArea,
          subAreaValues: subAreaValues,
          subAreaCount: subAreaValues.length
        });
      }

      return {
        isValid: allValid,
        totalArea: totalArea,
        perArea: perArea
      };
    }

    function renderPerAreaResults(perAreaResults) {
      var html = '';
      var i;
      for (i = 0; i < perAreaResults.length; i++) {
        var row = perAreaResults[i];
        var verticalSizeText = (row.vertical !== null) ? (row.vertical.pipe.size + '"') : 'N/A';
        var verticalCapText = (row.vertical !== null) ? (formatNumber(row.vertical.pipe.gpm, 0) + ' GPM') : 'N/A';
        var subAreaText = (row.subAreaCount > 0) ? row.subAreaListText + ' sq ft' : 'None';
        var runoffText = formatNumber(row.effectiveGpm, 2) + ' GPM / ' + formatNumber(row.effectiveCfs, 3) + ' CFS';
        html += '<div class="storm-area-card">';
        html += '  <div class="storm-area-header">' + row.label + '</div>';
        html += '  <div class="storm-area-inputs">';
        html += '    <div class="storm-area-input-item"><span class="storm-area-label">Main Area</span><span class="storm-area-value">' + formatNumber(row.parentArea, 0) + ' sq ft</span></div>';
        html += '    <div class="storm-area-input-item"><span class="storm-area-label">Sub Areas Total</span><span class="storm-area-value">' + subAreaText + '</span></div>';
        html += '    <div class="storm-area-input-item"><span class="storm-area-label">Side Wall Sum</span><span class="storm-area-value">' + formatNumber(row.sideWallSum, 0) + ' sq ft</span></div>';
        html += '  </div>';
        html += '  <div class="storm-area-key-results">';
        html += '    <div class="storm-area-key-item"><span class="storm-area-label">Effective Drainage Area</span><span class="storm-area-value">' + formatNumber(row.effectiveDrainageArea, 0) + ' sq ft</span></div>';
        html += '    <div class="storm-area-key-item"><span class="storm-area-label">Runoff</span><span class="storm-area-value">' + runoffText + '</span></div>';
        html += '    <div class="storm-area-key-item"><span class="storm-area-label">Horizontal Drain Size</span><span class="storm-area-value storm-size-value">' + row.horizontal.pipe.size + '"</span></div>';
        html += '    <div class="storm-area-key-item"><span class="storm-area-label">Vertical Drain Size</span><span class="storm-area-value storm-size-value">' + verticalSizeText + '</span></div>';
        html += '  </div>';
        html += '  <div class="storm-area-capacity">';
        html += '    <div class="storm-area-cap-item"><span class="storm-area-label">Maximum Horizontal Capacity</span><span class="storm-area-value">' + formatNumber(row.horizontal.pipe.gpm, 0) + ' GPM</span></div>';
        html += '    <div class="storm-area-cap-item"><span class="storm-area-label">Maximum Vertical Capacity</span><span class="storm-area-value">' + verticalCapText + '</span></div>';
        html += '  </div>';
        html += '  <div class="storm-recommendation-strip"><span class="storm-recommendation-label">Final Pipe Sizing</span><span class="storm-recommendation-values"><span>H: ' + row.horizontal.pipe.size + '"</span><span>V: ' + verticalSizeText + '</span></span></div>';
        html += '  <div class="storm-area-support">';
        html += '    <div class="storm-area-support-item"><span class="storm-area-label">Applied Side Wall Factor</span><span class="storm-area-value">' + formatNumber(row.sideWallFactor * 100, 0) + '%</span></div>';
        html += '    <div class="storm-area-support-item"><span class="storm-area-label">Side Wall Contribution</span><span class="storm-area-value">' + formatNumber(row.sideWallContribution, 0) + ' sq ft</span></div>';
        html += '  </div>';
        if (row.demandMultiplier > 1) {
          html += '  <div class="storm-area-note">Runoff base: ' + formatNumber(row.baseGpm, 2) + ' GPM (' + formatNumber(row.baseCfs, 3) + ' CFS). Effective uses ' + formatNumber(row.demandMultiplier, 2) + 'x demand.</div>';
        }
        if (row.warningText !== '') {
          html += '<div class="warning-line">' + row.warningText + '</div>';
        }
        html += '</div>';
      }
      document.getElementById('perAreaResultsList').innerHTML = html;
    }

    function clearValidationSummary() {
      document.getElementById('validationSummary').innerHTML = '';
      clearInputError(document.getElementById('rainfallRate'), null);
    }

    function showValidationSummary(message) {
      document.getElementById('validationSummary').innerHTML = message;
    }

    function updateSlopeOptions() {
      var codeBasis = document.getElementById('codeBasis').value;
      var slopeSelect = document.getElementById('slope');
      var existing = slopeSelect.value;
      slopeSelect.options.length = 0;

      var key, opt, j;
      for (key in pipeData[codeBasis].horizontal) {
        opt = document.createElement('option');
        opt.value = key;
        if (key == "1/16 in/ft") opt.text = '1/16 inch per foot';
        else if (key == "1/8 in/ft") opt.text = '1/8 inch per foot';
        else if (key == "1/4 in/ft") opt.text = '1/4 inch per foot';
        else if (key == "1/2 in/ft") opt.text = '1/2 inch per foot';
        else opt.text = key;
        slopeSelect.add(opt);
      }

      for (j = 0; j < slopeSelect.options.length; j++) {
        if (slopeSelect.options[j].value == existing) {
          slopeSelect.selectedIndex = j;
          return;
        }
      }
      if (slopeSelect.options.length > 0) slopeSelect.selectedIndex = 0;
    }

    function setZipLookupStatus(message, statusClass) {
      var node = document.getElementById('zipLookupStatus');
      if (!node) return;
      node.className = 'lookup-status';
      if (statusClass) node.className += ' ' + statusClass;
      node.innerHTML = message || '';
    }

    async function lookupRainfallByZip() {
      var zipInput = document.getElementById('zipCode');
      var rainfallInput = document.getElementById('rainfallRate');
      var zipCode = zipInput ? zipInput.value : '';
      if (!zipCode || zipCode.replace(/\D/g, '').length === 0) {
        setZipLookupStatus('Enter a ZIP code to find rainfall.', 'warn');
        return;
      }
      if (!window.pywebview || !window.pywebview.api || !window.pywebview.api.lookup_zip_rainfall) {
        setZipLookupStatus('ZIP lookup service is not ready yet. Please try again.', 'warn');
        return;
      }
      setZipLookupStatus('Looking up ZIP rainfall data...', '');
      try {
        var result = await window.pywebview.api.lookup_zip_rainfall(zipCode);
        if (!result || !result.ok) {
          setZipLookupStatus((result && result.reason) ? result.reason : 'ZIP lookup failed.', 'warn');
          return;
        }
        if (rainfallInput) rainfallInput.value = parseFloat(result.rainfall_in_per_hr).toFixed(1);
        if (zipInput) zipInput.value = result.zip;
        setZipLookupStatus(
          'ZIP ' + result.zip + ' (' + result.city + ', ' + result.state_id + '): ' +
          'raw ' + parseFloat(result.rainfall_source_in_per_hr).toFixed(3) + ' in/hr -> ' +
          'used ' + parseFloat(result.rainfall_in_per_hr).toFixed(1) + ' in/hr',
          'ok'
        );
        recalculate();
      } catch (e) {
        setZipLookupStatus('ZIP lookup failed. ' + ((e && e.message) ? e.message : 'Unknown error'), 'warn');
      }
    }

    function syncColumnHeights() {
      return;
    }

    function parseNonNegativeNumber(rawValue, fallback) {
      var n = parseFloat(rawValue);
      if (isNaN(n) || n < 0) return fallback;
      return n;
    }

    function isPositiveNumber(rawValue) {
      return !!(window.Utils && window.Utils.isPositiveNumber ? window.Utils.isPositiveNumber(rawValue) : (!isNaN(parseFloat(rawValue)) && parseFloat(rawValue) > 0));
    }

    function isNonNegativeNumber(rawValue) {
      return !!(window.Utils && window.Utils.isNonNegativeNumber ? window.Utils.isNonNegativeNumber(rawValue) : (!isNaN(parseFloat(rawValue)) && parseFloat(rawValue) >= 0));
    }

    function setWarningResult(node, message) {
      if (window.Utils && window.Utils.setWarningHtml) {
        window.Utils.setWarningHtml(node, message);
        return;
      }
      if (node) node.innerHTML = '<div class="warning-line">' + message + '</div>';
    }

    function setPlaceholderResult(node, message) {
      if (window.Utils && window.Utils.setPlaceholderHtml) {
        window.Utils.setPlaceholderHtml(node, message);
        return;
      }
      if (node) node.innerHTML = '<div class="result-placeholder">' + message + '</div>';
    }

    function blankProjectModel() {
      return {
        projectName: '',
        projectNumber: '',
        clientName: '',
        projectLocation: '',
        engineerName: '',
        companyName: '',
        date: '',
        revision: '',
        notes: ''
      };
    }

    function normalizeProjectInfo(raw) {
      var source = (raw && typeof raw === 'object') ? raw : {};
      var model = blankProjectModel();
      model.projectName = String(source.projectName || '').trim();
      model.projectNumber = String(source.projectNumber || '').trim();
      model.clientName = String(source.clientName || '').trim();
      model.projectLocation = String(source.projectLocation || '').trim();
      model.engineerName = String(source.engineerName || '').trim();
      model.companyName = String(source.companyName || '').trim();
      model.date = String(source.date || '').trim();
      model.revision = String(source.revision || '').trim();
      model.notes = String(source.notes || '').trim();
      return model;
    }

    function hasActiveProject() {
      return !!(activeProject && activeProject.projectName && activeProject.projectNumber && activeProject.clientName && activeProject.projectLocation);
    }

    function getProjectContextPayload() {
      return {
        activeProject: hasActiveProject() ? normalizeProjectInfo(activeProject) : null,
        savedProjects: savedProjects.map(function (item) { return normalizeProjectInfo(item); })
      };
    }

    function upsertSavedProject(project) {
      var p = normalizeProjectInfo(project);
      if (!p.projectName) return;
      var key = (p.projectNumber + '|' + p.projectName).toLowerCase();
      var i;
      for (i = 0; i < savedProjects.length; i++) {
        var existing = savedProjects[i];
        var existingKey = ((existing.projectNumber || '') + '|' + (existing.projectName || '')).toLowerCase();
        if (existingKey === key) {
          savedProjects[i] = p;
          return;
        }
      }
      savedProjects.push(p);
    }

    function setActiveProject(project) {
      activeProject = normalizeProjectInfo(project);
      if (hasActiveProject()) upsertSavedProject(activeProject);
      if (window.updateAppState) {
        window.updateAppState('activeProject', activeProject);
        window.updateAppState('savedProjects', savedProjects.slice());
      }
      renderCurrentProjectCard();
    }

    function projectFormValue(id) {
      var node = document.getElementById(id);
      return node ? String(node.value || '').trim() : '';
    }

    function fillProjectForm(project) {
      var data = normalizeProjectInfo(project);
      if (document.getElementById('projectNameInput')) document.getElementById('projectNameInput').value = data.projectName;
      if (document.getElementById('projectNumberInput')) document.getElementById('projectNumberInput').value = data.projectNumber;
      if (document.getElementById('projectClientInput')) document.getElementById('projectClientInput').value = data.clientName;
      if (document.getElementById('projectLocationInput')) document.getElementById('projectLocationInput').value = data.projectLocation;
      if (document.getElementById('projectEngineerInput')) document.getElementById('projectEngineerInput').value = data.engineerName;
      if (document.getElementById('projectCompanyInput')) document.getElementById('projectCompanyInput').value = data.companyName;
      if (document.getElementById('projectDateInput')) document.getElementById('projectDateInput').value = data.date;
      if (document.getElementById('projectRevisionInput')) document.getElementById('projectRevisionInput').value = data.revision;
      if (document.getElementById('projectNotesInput')) document.getElementById('projectNotesInput').value = data.notes;
    }

    function openProjectForm(mode) {
      projectEditMode = mode || (hasActiveProject() ? 'edit' : 'create');
      var modal = document.getElementById('projectFormModal');
      var titleNode = document.getElementById('projectFormTitle');
      var errorNode = document.getElementById('projectFormError');
      if (errorNode) {
        errorNode.style.display = 'none';
        errorNode.innerHTML = '';
      }
      if (titleNode) titleNode.innerHTML = (projectEditMode === 'edit') ? 'Edit Project Information' : 'Create Project Information';
      fillProjectForm(projectEditMode === 'edit' ? activeProject : blankProjectModel());
      if (modal) modal.style.display = 'flex';
    }

    function closeProjectForm() {
      var modal = document.getElementById('projectFormModal');
      if (modal) modal.style.display = 'none';
    }

    function saveProjectFromForm() {
      var project = normalizeProjectInfo({
        projectName: projectFormValue('projectNameInput'),
        projectNumber: projectFormValue('projectNumberInput'),
        clientName: projectFormValue('projectClientInput'),
        projectLocation: projectFormValue('projectLocationInput'),
        engineerName: projectFormValue('projectEngineerInput'),
        companyName: projectFormValue('projectCompanyInput'),
        date: projectFormValue('projectDateInput'),
        revision: projectFormValue('projectRevisionInput'),
        notes: projectFormValue('projectNotesInput')
      });
      var errorNode = document.getElementById('projectFormError');
      if (!project.projectName || !project.projectNumber || !project.clientName || !project.projectLocation) {
        if (errorNode) {
          errorNode.style.display = 'block';
          errorNode.innerHTML = 'Project Name, Project Number, Client Name, and Project Location are required.';
        }
        return;
      }
      setActiveProject(project);
      closeProjectForm();
    }

    function openProjectDetails() {
      if (!hasActiveProject()) {
        openProjectForm('create');
        return;
      }
      var modal = document.getElementById('projectDetailsModal');
      var body = document.getElementById('projectDetailsBody');
      if (body) {
        var p = normalizeProjectInfo(activeProject);
        body.innerHTML =
          '<div><div class="project-context-item-label">Project Name</div><div class="project-context-item-value">' + p.projectName + '</div></div>' +
          '<div><div class="project-context-item-label">Project Number</div><div class="project-context-item-value">' + p.projectNumber + '</div></div>' +
          '<div><div class="project-context-item-label">Client</div><div class="project-context-item-value">' + p.clientName + '</div></div>' +
          '<div><div class="project-context-item-label">Location</div><div class="project-context-item-value">' + p.projectLocation + '</div></div>' +
          '<div><div class="project-context-item-label">Engineer</div><div class="project-context-item-value">' + (p.engineerName || '-') + '</div></div>' +
          '<div><div class="project-context-item-label">Company</div><div class="project-context-item-value">' + (p.companyName || '-') + '</div></div>' +
          '<div><div class="project-context-item-label">Date</div><div class="project-context-item-value">' + (p.date || '-') + '</div></div>' +
          '<div><div class="project-context-item-label">Revision</div><div class="project-context-item-value">' + (p.revision || '-') + '</div></div>' +
          '<div class="project-details-full"><div class="project-context-item-label">Notes</div><div class="project-context-item-value">' + (p.notes || '-') + '</div></div>';
      }
      if (modal) modal.style.display = 'flex';
    }

    function closeProjectDetails() {
      var modal = document.getElementById('projectDetailsModal');
      if (modal) modal.style.display = 'none';
    }

    function editProjectFromDetails() {
      closeProjectDetails();
      openProjectForm('edit');
    }

    function setProjectAsActiveFromDetails() {
      if (!hasActiveProject()) return;
      setActiveProject(activeProject);
      closeProjectDetails();
    }

    function continueWithoutProject() {
      renderCurrentProjectCard();
    }

    function setHomeVersionLabel(versionText) {
      var homeVersionNode = document.getElementById('homeVersionLabel');
      if (!homeVersionNode) return;
      if (versionText) {
        homeVersionNode.textContent = versionText;
        homeVersionNode.style.visibility = 'visible';
      } else {
        // Keep left footer slot reserved so the right credit stays right-aligned.
        homeVersionNode.textContent = '';
        homeVersionNode.style.visibility = 'hidden';
      }
      homeVersionNode.style.display = 'block';
    }

    function renderCurrentProjectCard() {
      var body = document.getElementById('projectContextBody');
      var versionNode = document.getElementById('projectContextVersion');
      var versionText = '';
      if (APP_VERSION) {
        versionText = APP_VERSION.charAt(0).toLowerCase() === 'v' ? APP_VERSION : ('v' + APP_VERSION);
      }
      if (versionNode) versionNode.innerHTML = '';
      setHomeVersionLabel(versionText);
      if (!body) return;
      if (!hasActiveProject()) {
        body.innerHTML =
          '<div class="project-context-main">' +
          '<div class="project-context-title">No active project selected</div>' +
          '<div class="project-context-subtle">Calculations can still be performed. Saved/exported reports will not include project data.</div>' +
          '</div>' +
          '<div class="project-context-actions">' +
          '<input type="button" class="btn project-context-btn-primary" value="CREATE PROJECT" onclick="openProjectForm(\'create\');" />' +
          '<input type="button" class="btn project-context-btn-secondary" value="CONTINUE WITHOUT PROJECT" onclick="continueWithoutProject();" />' +
          '</div>';
        return;
      }
      var p = normalizeProjectInfo(activeProject);
      body.innerHTML =
        '<div class="project-context-main">' +
        '<div class="project-context-line"><strong>Project:</strong> ' + p.projectName + ' | <strong>No:</strong> ' + p.projectNumber + ' | <strong>Client:</strong> ' + p.clientName + ' | ' + p.projectLocation + '</div>' +
        '<div class="project-context-subtle">Revision: ' + (p.revision || '-') + ' | Date: ' + (p.date || '-') + '</div>' +
        '</div>' +
        '<div class="project-context-actions">' +
        '<input type="button" class="btn project-context-btn-secondary" value="NEW PROJECT" onclick="openProjectForm(\'create\');" />' +
        '<input type="button" class="btn project-context-btn-secondary" value="EDIT INFO" onclick="openProjectForm(\'edit\');" />' +
        '<input type="button" class="btn project-context-btn-secondary" value="VIEW DETAILS" onclick="openProjectDetails();" />' +
        '</div>';
    }

    function confirmProjectContextForAction(actionLabel) {
      if (hasActiveProject()) return true;
      var proceed = window.confirm('This ' + actionLabel + ' is not linked to a project. Press OK to continue without project information, or Cancel to create a project first.');
      if (!proceed) {
        openProjectForm('create');
        return false;
      }
      return true;
    }

    function formatCriteriaValue(value) {
      var n = parseFloat(value);
      if (isNaN(n)) return '';
      var text = n.toFixed(4);
      text = text.replace(/0+$/, '').replace(/\.$/, '');
      return text;
    }

    function getDistinctSortedNumbers(values) {
      var seen = {};
      var list = [];
      var i, raw, key;
      for (i = 0; i < values.length; i++) {
        raw = parseFloat(values[i]);
        if (isNaN(raw)) continue;
        key = formatCriteriaValue(raw);
        if (!key || seen[key]) continue;
        seen[key] = true;
        list.push(raw);
      }
      list.sort(function (a, b) { return a - b; });
      return list;
    }

    function fillNumericSelect(selectId, values, selectedValue) {
      var node = document.getElementById(selectId);
      if (!node) return '';
      node.options.length = 0;
      var i, opt, text, selectedText = formatCriteriaValue(selectedValue);
      for (i = 0; i < values.length; i++) {
        text = formatCriteriaValue(values[i]);
        opt = document.createElement('option');
        opt.value = text;
        opt.text = text;
        node.add(opt);
      }
      if (node.options.length <= 0) return '';
      for (i = 0; i < node.options.length; i++) {
        if (node.options[i].value === selectedText) {
          node.selectedIndex = i;
          return node.options[i].value;
        }
      }
      node.selectedIndex = 0;
      return node.options[0].value;
    }

    function getGasCriteriaCandidates(codeBasis, gasType, material) {
      var out = [];
      var i;
      for (i = 0; i < gasTableCatalog.length; i++) {
        if (gasTableCatalog[i].codeBasis === codeBasis && gasTableCatalog[i].gasType === gasType && gasTableCatalog[i].material === material) {
          out.push(gasTableCatalog[i]);
        }
      }
      return out;
    }

    function refreshGasCriteriaDropdowns() {
      var codeBasisNode = document.getElementById('gasCodeBasis');
      var gasTypeNode = document.getElementById('gasGasType');
      var materialNode = document.getElementById('gasMaterial');
      var inletNode = document.getElementById('gasInletPressure');
      if (!codeBasisNode || !gasTypeNode || !materialNode) return;

      var codeBasis = codeBasisNode.value;
      var gasType = gasTypeNode.value;
      var material = materialNode.value || 'SCHEDULE_40_METALLIC';
      var baseCandidates = getGasCriteriaCandidates(codeBasis, gasType, material);
      if (baseCandidates.length <= 0) return;

      var currentInlet = inletNode ? inletNode.value : '';
      var inletValues = getDistinctSortedNumbers(baseCandidates.map(function (t) { return t.inletPressure; }));
      var selectedInletText = '';
      var selectedInlet = 0;
      var i;

      if (inletNode) {
        inletNode.options.length = 0;
        var lt2Option = document.createElement('option');
        lt2Option.value = 'LT2';
        lt2Option.text = 'Less Than 2';
        inletNode.add(lt2Option);
        for (i = 0; i < inletValues.length; i++) {
          var inletOption = document.createElement('option');
          inletOption.value = formatCriteriaValue(inletValues[i]);
          inletOption.text = formatCriteriaValue(inletValues[i]);
          inletNode.add(inletOption);
        }

        if (currentInlet === 'LT2') {
          inletNode.value = 'LT2';
          selectedInletText = 'LT2';
        } else if (currentInlet) {
          inletNode.value = formatCriteriaValue(currentInlet);
          selectedInletText = inletNode.value || '';
          if (!selectedInletText && inletNode.options.length > 1) {
            inletNode.selectedIndex = 1;
            selectedInletText = inletNode.options[1].value;
          }
        } else if (inletNode.options.length > 1) {
          inletNode.selectedIndex = 1;
          selectedInletText = inletNode.options[1].value;
        } else {
          inletNode.selectedIndex = 0;
          selectedInletText = inletNode.options[0].value;
        }
      }

      if (selectedInletText === 'LT2') selectedInlet = inletValues.length > 0 ? inletValues[0] : 0;
      else selectedInlet = parseFloat(selectedInletText);

      var dropCandidates = [];
      for (i = 0; i < baseCandidates.length; i++) {
        if (Math.abs(parseNonNegativeNumber(baseCandidates[i].inletPressure, 0) - selectedInlet) < 0.0001) dropCandidates.push(baseCandidates[i]);
      }
      if (dropCandidates.length <= 0) dropCandidates = baseCandidates;

      var currentDrop = document.getElementById('gasPressureDrop') ? document.getElementById('gasPressureDrop').value : '';
      var dropValues = getDistinctSortedNumbers(dropCandidates.map(function (t) { return t.pressureDrop; }));
      var selectedDropText = fillNumericSelect('gasPressureDrop', dropValues, currentDrop);
      var selectedDrop = parseFloat(selectedDropText);

      var sgCandidates = [];
      for (i = 0; i < dropCandidates.length; i++) {
        if (Math.abs(parseNonNegativeNumber(dropCandidates[i].pressureDrop, 0) - selectedDrop) < 0.0001) sgCandidates.push(dropCandidates[i]);
      }
      if (sgCandidates.length <= 0) sgCandidates = dropCandidates;

      var currentSg = document.getElementById('gasSpecificGravity') ? document.getElementById('gasSpecificGravity').value : '';
      var sgValues = getDistinctSortedNumbers(sgCandidates.map(function (t) { return t.specificGravity; }));
      fillNumericSelect('gasSpecificGravity', sgValues, currentSg);
      calculateGasPipe(false);
    }

    function initGasMaterialOptions() {
      var node = document.getElementById('gasMaterial');
      if (!node) return;
      node.options.length = 0;
      var i, opt;
      for (i = 0; i < GAS_MATERIAL_OPTIONS.length; i++) {
        opt = document.createElement('option');
        opt.value = GAS_MATERIAL_OPTIONS[i].id;
        opt.text = GAS_MATERIAL_OPTIONS[i].label;
        node.add(opt);
      }
      node.value = 'SCHEDULE_40_METALLIC';
    }

    function getGasSelectionState() {
      var demandMode = document.getElementById('gasDemandMode') ? document.getElementById('gasDemandMode').value : 'CFH';
      var heatingValue = parseNonNegativeNumber(document.getElementById('gasHeatingValue') ? document.getElementById('gasHeatingValue').value : 0, 0);
      var inletRaw = document.getElementById('gasInletPressure') ? document.getElementById('gasInletPressure').value : '';
      var inletOption = (inletRaw === 'LT2') ? 'LT2' : '';
      return {
        codeBasis: (document.getElementById('gasCodeBasis') ? document.getElementById('gasCodeBasis').value : 'IPC'),
        gasType: (document.getElementById('gasGasType') ? document.getElementById('gasGasType').value : 'NATURAL_GAS'),
        inletPressureOption: inletOption,
        inletPressure: parseNonNegativeNumber((inletOption === 'LT2') ? 0 : inletRaw, 0),
        pressureDrop: parseNonNegativeNumber(document.getElementById('gasPressureDrop') ? document.getElementById('gasPressureDrop').value : 0, 0),
        specificGravity: parseNonNegativeNumber(document.getElementById('gasSpecificGravity') ? document.getElementById('gasSpecificGravity').value : 0, 0),
        material: (document.getElementById('gasMaterial') ? document.getElementById('gasMaterial').value : 'SCHEDULE_40_METALLIC'),
        demandMode: demandMode,
        heatingValue: heatingValue
      };
    }

    function getGasDemandCfh(selection, rawDemandValue) {
      var demandValue = parseNonNegativeNumber(rawDemandValue, 0);
      if (selection.demandMode === 'BTUH') {
        if (selection.heatingValue <= 0) return 0;
        return demandValue / selection.heatingValue;
      }
      return demandValue;
    }

    function makeGasLineId() {
      return 'line_' + (new Date().getTime()) + '_' + Math.floor(Math.random() * 1000);
    }

    function addGasLineRow(line) {
      var list = document.getElementById('gasLinesBody');
      if (!list) return;
      var lineObj = line || { id: makeGasLineId(), label: '', demandValue: 0, runLength: 0 };
      var html = '';
      html += '<tr class="gas-line-row" data-line-id="' + lineObj.id + '">';
      html += '  <td><input type="text" class="gas-line-label" value="' + (lineObj.label || '') + '" placeholder="Line 1, Branch A, Unit 3..." oninput="calculateGasPipe(false);" onchange="calculateGasPipe(false);" /></td>';
      html += '  <td class="num"><input type="number" step="0.01" class="gas-line-demand" value="' + parseNonNegativeNumber(lineObj.demandValue, 0) + '" oninput="calculateGasPipe(false);" onchange="calculateGasPipe(false);" /></td>';
      html += '  <td class="num"><input type="number" step="0.01" class="gas-line-length" value="' + parseNonNegativeNumber(lineObj.runLength, 0) + '" oninput="calculateGasPipe(false);" onchange="calculateGasPipe(false);" /></td>';
      html += '  <td class="num"><input type="button" class="btn btn-small gas-remove-btn" value="Remove" onclick="removeGasLineRow(this);" /></td>';
      html += '</tr>';
      list.insertAdjacentHTML('beforeend', html);
      calculateGasPipe(false);
    }

    function removeGasLineRow(btn) {
      var list = document.getElementById('gasLinesBody');
      if (!list) return;
      var rows = list.getElementsByClassName('gas-line-row');
      if (rows.length <= 1) {
        alert('At least one gas line is required.');
        return;
      }
      var row = btn;
      while (row && (' ' + row.className + ' ').indexOf(' gas-line-row ') < 0) row = row.parentNode;
      if (row && row.parentNode) row.parentNode.removeChild(row);
      calculateGasPipe(false);
    }

    function clearGasLineRows() {
      var list = document.getElementById('gasLinesBody');
      if (list) list.innerHTML = '';
    }

    function getGasLinesState() {
      var list = document.getElementById('gasLinesBody');
      var rows = list ? list.getElementsByClassName('gas-line-row') : [];
      var lines = [];
      var i;
      for (i = 0; i < rows.length; i++) {
        lines.push({
          id: rows[i].getAttribute('data-line-id') || makeGasLineId(),
          label: (rows[i].getElementsByClassName('gas-line-label')[0] || { value: '' }).value || ('Line ' + (i + 1)),
          demandValue: parseNonNegativeNumber((rows[i].getElementsByClassName('gas-line-demand')[0] || { value: 0 }).value, 0),
          runLength: parseNonNegativeNumber((rows[i].getElementsByClassName('gas-line-length')[0] || { value: 0 }).value, 0)
        });
      }
      return lines;
    }

    function getConservativeTableCandidate(selection) {
      var candidates = [];
      var i;
      for (i = 0; i < gasTableCatalog.length; i++) {
        if (gasTableCatalog[i].codeBasis === selection.codeBasis && gasTableCatalog[i].gasType === selection.gasType && gasTableCatalog[i].material === selection.material) {
          candidates.push(gasTableCatalog[i]);
        }
      }
      if (candidates.length <= 0) return null;

      var best = null;
      var bestScore = 999999999;
      var score, inletDiff, dropDiff, sgDiff;
      for (i = 0; i < candidates.length; i++) {
        inletDiff = candidates[i].inletPressure - selection.inletPressure;
        dropDiff = candidates[i].pressureDrop - selection.pressureDrop;
        sgDiff = candidates[i].specificGravity - selection.specificGravity;
        score = 0;

        if (inletDiff > 0) score += 300 + inletDiff * 20; else score += Math.abs(inletDiff) * 4;
        if (dropDiff > 0) score += 300 + dropDiff * 200; else score += Math.abs(dropDiff) * 40;
        if (sgDiff < 0) score += 300 + Math.abs(sgDiff) * 500; else score += Math.abs(sgDiff) * 80;

        if (score < bestScore) {
          bestScore = score;
          best = candidates[i];
        }
      }
      return best;
    }

    function getExactTableCandidates(selection) {
      var filtered = [];
      var i;
      for (i = 0; i < gasTableCatalog.length; i++) {
        if (
          gasTableCatalog[i].codeBasis === selection.codeBasis &&
          gasTableCatalog[i].gasType === selection.gasType &&
          gasTableCatalog[i].material === selection.material &&
          Math.abs(gasTableCatalog[i].inletPressure - selection.inletPressure) < 0.0001 &&
          Math.abs(gasTableCatalog[i].pressureDrop - selection.pressureDrop) < 0.0001 &&
          Math.abs(gasTableCatalog[i].specificGravity - selection.specificGravity) < 0.0001
        ) {
          filtered.push(gasTableCatalog[i]);
        }
      }
      return filtered;
    }

    function calculateGasPipe(commitResults) {
      if (commitResults !== true) commitResults = false;
      var selection = getGasSelectionState();
      var effectiveSelection = {
        codeBasis: selection.codeBasis,
        gasType: selection.gasType,
        material: selection.material,
        inletPressure: selection.inletPressure,
        pressureDrop: selection.pressureDrop,
        specificGravity: selection.specificGravity
      };
      var lines = getGasLinesState();
      var resultNode = document.getElementById('gasResults');
      var previewNode = document.getElementById('gasPreview');
      var selectedTable = null;
      var baseCandidates = getGasCriteriaCandidates(selection.codeBasis, selection.gasType, selection.material);
      var minInlet = 0;
      var inletResolutionNote = '';
      var exactCandidates;
      var conservative = null;
      var matchType = '';
      var warning = '';
      var i, li;

      if (lines.length <= 0) {
        if (previewNode) previewNode.innerHTML = 'Add at least one gas line.';
        if (commitResults) {
          latestGasSnapshot = null;
          setWarningResult(resultNode, 'Add at least one gas line before calculating.');
        }
        return;
      }

      if (selection.inletPressureOption === 'LT2') {
        if (baseCandidates.length <= 0) {
          if (previewNode) previewNode.innerHTML = 'No internal gas table is available for this selection.';
          if (commitResults) {
            latestGasSnapshot = null;
            setWarningResult(resultNode, 'No internal gas table is available for this selection.');
          }
          return;
        }
        minInlet = parseNonNegativeNumber(baseCandidates[0].inletPressure, 0);
        for (i = 1; i < baseCandidates.length; i++) {
          if (parseNonNegativeNumber(baseCandidates[i].inletPressure, 0) < minInlet) minInlet = parseNonNegativeNumber(baseCandidates[i].inletPressure, 0);
        }
        effectiveSelection.inletPressure = minInlet;
        inletResolutionNote = '"Less Than 2" mapped to lowest available inlet pressure table: ' + formatCriteriaValue(minInlet) + '.';
      }

      exactCandidates = getExactTableCandidates(effectiveSelection);
      if (exactCandidates.length > 0) {
        selectedTable = exactCandidates[0];
        matchType = 'exact';
      } else {
        conservative = getConservativeTableCandidate(effectiveSelection);
        selectedTable = conservative;
        matchType = conservative ? 'conservative_fallback' : '';
      }

      if (!selectedTable) {
        if (previewNode) previewNode.innerHTML = 'No internal gas table is available for this selection.';
        if (commitResults) {
          latestGasSnapshot = null;
          setWarningResult(resultNode, 'No internal gas table is available for this selection.');
        }
        return;
      }

      if (matchType === 'conservative_fallback') {
        warning = 'No exact criteria match found. Conservative fallback table was used: ' + selectedTable.tableName + '.';
      }
      if (inletResolutionNote) {
        warning = warning ? (warning + ' ' + inletResolutionNote) : inletResolutionNote;
      }

      var rows = matrixToRows(selectedTable);
      if (rows.length <= 0) {
        if (previewNode) previewNode.innerHTML = 'Selected gas table has no usable rows.';
        if (commitResults) {
          latestGasSnapshot = null;
          setWarningResult(resultNode, 'Selected gas table has no usable rows.');
        }
        return;
      }

      var lineResults = [];
      for (li = 0; li < lines.length; li++) {
        var demandCfh = getGasDemandCfh(selection, lines[li].demandValue);
        var runLength = parseNonNegativeNumber(lines[li].runLength, 0);
        if (demandCfh <= 0 || runLength <= 0) {
          lineResults.push({
            lineId: lines[li].id,
            lineLabel: lines[li].label || ('Line ' + (li + 1)),
            demandCfh: demandCfh,
            runLength: runLength,
            recommendedSize: 'N/A',
            capacityCfh: 0,
            selectedLength: 0,
            maxAllowableLength: 0,
            warning: 'Invalid demand or run length.'
          });
          continue;
        }

        var selectedRow = rows[rows.length - 1];
        for (i = 0; i < rows.length; i++) {
          if (runLength <= rows[i].length) {
            selectedRow = rows[i];
            break;
          }
        }

        var chosenSize = '';
        var chosenCapacity = 0;
        for (i = 0; i < GAS_PIPE_SIZES.length; i++) {
          if (parseNonNegativeNumber(selectedRow.capacities[GAS_PIPE_SIZES[i]], 0) >= demandCfh) {
            chosenSize = GAS_PIPE_SIZES[i];
            chosenCapacity = parseNonNegativeNumber(selectedRow.capacities[GAS_PIPE_SIZES[i]], 0);
            break;
          }
        }

        if (chosenSize === '') {
          lineResults.push({
            lineId: lines[li].id,
            lineLabel: lines[li].label || ('Line ' + (li + 1)),
            demandCfh: demandCfh,
            runLength: runLength,
            recommendedSize: 'OUT OF RANGE',
            capacityCfh: parseNonNegativeNumber(selectedRow.capacities[GAS_PIPE_SIZES[GAS_PIPE_SIZES.length - 1]], 0),
            selectedLength: selectedRow.length,
            maxAllowableLength: 0,
            warning: 'Demand exceeds table capacity at selected length.'
          });
          continue;
        }

        var maxAllowableLength = 0;
        for (i = 0; i < rows.length; i++) {
          if (parseNonNegativeNumber(rows[i].capacities[chosenSize], 0) >= demandCfh) {
            maxAllowableLength = rows[i].length;
          }
        }

        lineResults.push({
          lineId: lines[li].id,
          lineLabel: lines[li].label || ('Line ' + (li + 1)),
          demandCfh: demandCfh,
          runLength: runLength,
          recommendedSize: chosenSize,
          capacityCfh: chosenCapacity,
          selectedLength: selectedRow.length,
          maxAllowableLength: maxAllowableLength,
          warning: ''
        });
      }

      var totalDemandCfh = 0;
      var maxSizeIndex = -1;
      var maxSizeLabel = 'N/A';
      for (li = 0; li < lineResults.length; li++) {
        totalDemandCfh += parseNonNegativeNumber(lineResults[li].demandCfh, 0);
        var currentIdx = GAS_PIPE_SIZES.indexOf(lineResults[li].recommendedSize);
        if (currentIdx > maxSizeIndex) {
          maxSizeIndex = currentIdx;
          maxSizeLabel = lineResults[li].recommendedSize;
        }
      }

      if (previewNode) {
        previewNode.innerHTML =
          '<div class="gas-preview-strip">' +
          '<span class="gas-preview-item"><span class="gas-preview-label">Code Basis</span><span class="gas-preview-value">' + selection.codeBasis + '</span></span>' +
          '<span class="gas-preview-item"><span class="gas-preview-label">Gas Type</span><span class="gas-preview-value">' + (selection.gasType === 'PROPANE' ? 'Propane' : 'Natural Gas') + '</span></span>' +
          '<span class="gas-preview-item"><span class="gas-preview-label">Material</span><span class="gas-preview-value">' + getGasMaterialLabel(selection.material) + '</span></span>' +
          '<span class="gas-preview-item"><span class="gas-preview-label">Total Demand</span><span class="gas-preview-value">' + formatNumber(totalDemandCfh, 2) + ' CFH</span></span>' +
          '<span class="gas-preview-item"><span class="gas-preview-label">Selected Pipe Size</span><span class="gas-preview-value">' + maxSizeLabel + '</span></span>' +
          '</div>';
      }

      if (!commitResults) return;

      latestGasSnapshot = {
        tableId: selectedTable.id,
        tableName: selectedTable.tableName,
        matchType: matchType,
        warning: warning,
        inletResolutionNote: inletResolutionNote,
        lines: lineResults
      };

      var html = '';
      if (warning) html += '<div class="warning-line">' + warning + '</div>';
      html += '<div class="gas-result-summary">';
      html += '<div class="gas-result-metric"><span class="gas-result-metric-label">Total Demand</span><span class="gas-result-metric-value">' + formatNumber(totalDemandCfh, 2) + ' CFH</span></div>';
      html += '<div class="gas-result-metric"><span class="gas-result-metric-label">Pipe Material</span><span class="gas-result-metric-value">' + getGasMaterialLabel(selection.material) + '</span></div>';
      html += '<div class="gas-result-metric"><span class="gas-result-metric-label">Sizing Condition</span><span class="gas-result-metric-value">' + selectedTable.tableName + '</span></div>';
      html += '<div class="gas-result-metric gas-result-metric-main"><span class="gas-result-metric-label">Largest Pipe Size</span><span class="gas-result-metric-value">' + maxSizeLabel + '</span></div>';
      html += '</div>';
      html += '<div class="gas-result-table-wrap"><table class="gas-result-table">';
      html += '<thead><tr><th>Line</th><th class="num">Demand</th><th class="num">Length</th><th class="num">Recommended Size</th><th class="num">Capacity</th><th class="num">Max Length</th><th>Warning</th></tr></thead><tbody>';
      for (li = 0; li < lineResults.length; li++) {
        html += '<tr>';
        html += '<td>' + lineResults[li].lineLabel + '</td>';
        html += '<td class="num">' + formatNumber(lineResults[li].demandCfh, 2) + '</td>';
        html += '<td class="num">' + formatNumber(lineResults[li].runLength, 1) + '</td>';
        html += '<td class="num gas-size-cell">' + lineResults[li].recommendedSize + '</td>';
        html += '<td class="num">' + formatNumber(lineResults[li].capacityCfh, 0) + '</td>';
        html += '<td class="num">' + (lineResults[li].maxAllowableLength > 0 ? formatNumber(lineResults[li].maxAllowableLength, 1) : 'N/A') + '</td>';
        html += '<td>' + (lineResults[li].warning ? ('<span class="gas-result-warning-cell">' + lineResults[li].warning + '</span>') : '') + '</td>';
        html += '</tr>';
      }
      html += '</tbody></table></div>';
      html += '<div class="gas-result-note">Selected internal table: ' + selectedTable.tableName + '. Values are based on internal schedule and selected criteria.</div>';
      resultNode.innerHTML = html;
    }

    function buildProjectPayload() {
      var blocks = getAreaBlocks();
      var areas = [];
      var i, j, block, mainInput, subRows, wallRows, subInput, wallInput;

      for (i = 0; i < blocks.length; i++) {
        block = blocks[i];
        mainInput = block.getElementsByClassName('area-main-input')[0];
        subRows = getSubAreaRows(block);
        wallRows = getSideWallRows(block);
        var subAreas = [];
        var sideWalls = [];

        for (j = 0; j < subRows.length; j++) {
          subInput = subRows[j].getElementsByClassName('subarea-input')[0];
          if (!subInput) continue;
          subAreas.push(parseNonNegativeNumber(subInput.value, 0));
        }
        for (j = 0; j < wallRows.length; j++) {
          wallInput = wallRows[j].getElementsByClassName('sidewall-input')[0];
          if (!wallInput) continue;
          sideWalls.push(parseNonNegativeNumber(wallInput.value, 0));
        }

        areas.push({
          name: getAreaNameFromBlock(block, i),
          mainArea: mainInput ? parseNonNegativeNumber(mainInput.value, 0) : 0,
          subAreas: subAreas,
          sideWalls: sideWalls
        });
      }

      if (areas.length <= 0) {
        areas.push({ name: getAreaLabel(0), mainArea: 0, subAreas: [], sideWalls: [] });
      }

      return {
        app_version: APP_VERSION,
        schema_version: APP_SCHEMA_VERSION,
        activeModule: activeModule,
        project: getProjectContextPayload(),
        settings: {
          codeBasis: document.getElementById('codeBasis').value,
          slope: document.getElementById('slope').value,
          zipCode: document.getElementById('zipCode').value,
          rainfallRate: parseNonNegativeNumber(document.getElementById('rainfallRate').value, APP_DEFAULTS.stormRainfallRate || 1),
          drainType: document.getElementById('drainType').value
        },
        areas: areas,
        condensate: {
          selection: {
            sizingMode: 'PER_UNIT_QTY',
            equipment: (condensateZones.length > 0 && condensateZones[0].equipmentRows.length > 0 ? condensateZones[0].equipmentRows[0].equipment : 'AHU'),
            tonsPerUnit: parseNonNegativeNumber((condensateZones.length > 0 && condensateZones[0].equipmentRows.length > 0 ? condensateZones[0].equipmentRows[0].tonsPerUnit : 0), 0),
            quantity: parseNonNegativeNumber((condensateZones.length > 0 && condensateZones[0].equipmentRows.length > 0 ? condensateZones[0].equipmentRows[0].quantity : (APP_DEFAULTS.condensateQty || 1)), APP_DEFAULTS.condensateQty || 1),
            totalTons: roundNumber(getCondensateZonesState().reduce(function (zoneSum, zoneObj) {
              var zoneRows = zoneObj && zoneObj.equipmentRows ? zoneObj.equipmentRows : [];
              return zoneSum + zoneRows.reduce(function (rowSum, rowObj) { return rowSum + parseNonNegativeNumber(rowObj.rowTotalTons, 0); }, 0);
            }, 0), 3),
            equipmentRows: (condensateZones.length > 0 ? getCondensateZonesState()[0].equipmentRows : []),
            zones: getCondensateZonesState()
          },
          result: (latestCondensateSnapshot && typeof latestCondensateSnapshot === 'object') ? latestCondensateSnapshot : {
            section: 'condensate',
            tableName: 'TABLE 814.3',
            sizingMode: 'PER_UNIT_QTY',
            equipment: 'AHU',
            equipmentLabel: 'Air Handler Unit (AHU)',
            tonsPerUnit: 0,
            quantity: APP_DEFAULTS.condensateQty || 1,
            totalTonsInput: 0,
            totalTonsUsed: 0,
            recommendedSize: '',
            warning: '',
            rows: [],
            zones: []
          }
        },
        vent: {
          zones: getSanitaryZonesState(),
          selection: {
            codeBasis: (document.getElementById('ventCodeBasis') ? document.getElementById('ventCodeBasis').value : 'IPC'),
            usageType: (document.getElementById('ventUsageType') && document.getElementById('ventUsageType').value === 'ASSEMBLY') ? 'ASSEMBLY' : ((document.getElementById('ventUsageType') && document.getElementById('ventUsageType').value === 'PUBLIC') ? 'PUBLIC' : 'PRIVATE'),
            drainageOrientation: (document.getElementById('ventDrainageOrientation') && document.getElementById('ventDrainageOrientation').value === 'VERTICAL') ? 'VERTICAL' : 'HORIZONTAL',
            useManualTotal: !!(document.getElementById('ventUseManualTotal') && document.getElementById('ventUseManualTotal').checked),
            autoTotalDfu: parseNonNegativeNumber(document.getElementById('ventAutoTotalDfu') ? document.getElementById('ventAutoTotalDfu').value : 0, 0),
            manualTotalDfu: parseNonNegativeNumber(document.getElementById('ventManualTotalDfu') ? document.getElementById('ventManualTotalDfu').value : 0, 0),
            totalDfu: parseNonNegativeNumber(document.getElementById('ventTotalDfu') ? document.getElementById('ventTotalDfu').value : 0, 0),
            fixtureRows: getVentFixtureRowsState(),
            zones: getSanitaryZonesState()
          },
          result: latestVentSnapshot || {
            section: 'vent',
            codeBasis: 'IPC',
            usageType: 'PRIVATE',
            drainageOrientation: 'HORIZONTAL',
            dfu: 0,
            recommendedSize: '',
            sizingTable: 'TABLE 703.2',
            tableMaxUnits: 0,
            maxLengthReference: '',
            warning: '',
            autoTotalDfu: 0,
            manualTotalDfu: 0,
            useManualTotal: false,
            rows: [],
            zones: []
          }
        },
        ventilation: {
          selection: getVentilationSelectionState(),
          result: (latestVentilationSnapshot && latestVentilationSnapshot.result) ? latestVentilationSnapshot.result : {
            status: 'incomplete',
            message: 'Select occupancy and enter zone data to calculate required outdoor air.',
            zones: []
          }
        },
        sanitary_drainage: {
          zones: getSanitaryZonesState()
        },
        gas: {
          tables: gasTableCatalog,
          selection: getGasSelectionState(),
          lines: getGasLinesState(),
          result: latestGasSnapshot || {
            tableId: '',
            tableName: '',
            matchType: '',
            warning: '',
            lines: []
          }
        },
        wsfu: {
          zones: getWsfuZonesState(),
          selection: {
            totalFu: parseNonNegativeNumber(document.getElementById('wsfuTotalFu') ? document.getElementById('wsfuTotalFu').value : 0, 0),
            flushType: (document.getElementById('wsfuFlushType') ? document.getElementById('wsfuFlushType').value : 'FLUSH_TANK'),
            designLengthFt: parseNonNegativeNumber(document.getElementById('wsfuDesignLength') ? document.getElementById('wsfuDesignLength').value : (APP_DEFAULTS.wsfuDesignLengthFt || 100), APP_DEFAULTS.wsfuDesignLengthFt || 100),
            usageType: (document.getElementById('wsfuUsageType') ? document.getElementById('wsfuUsageType').value : 'PRIVATE'),
            useManualTotal: !!(document.getElementById('wsfuUseManualTotal') && document.getElementById('wsfuUseManualTotal').checked),
            manualTotalFu: parseNonNegativeNumber(document.getElementById('wsfuManualTotalFu') ? document.getElementById('wsfuManualTotalFu').value : 0, 0),
            autoTotalFu: parseNonNegativeNumber(document.getElementById('wsfuAutoTotalFu') ? document.getElementById('wsfuAutoTotalFu').value : 0, 0),
            fixtureRows: getWsfuFixtureRowsState(),
            zones: getWsfuZonesState()
          },
          result: (latestWsfuSnapshot && latestWsfuSnapshot.result) ? latestWsfuSnapshot.result : {
            totalFu: 0,
            flushTankGpm: 0,
            flushValveGpm: 0,
            matchedTankFu: 0,
            matchedValveFu: 0,
            flushType: 'FLUSH_TANK',
            selectedFlowGpm: 0,
            designLengthFt: APP_DEFAULTS.wsfuDesignLengthFt || 100,
            serviceSize: '',
            sizeCapacityGpm: 0,
            frictionLossByPipeType: [],
            warning: '',
            source: ''
          }
        },
        fixtureUnit: {
          selection: getFixtureUnitSelectionState(),
          result: (latestFixtureUnitSnapshot && latestFixtureUnitSnapshot.result) ? latestFixtureUnitSnapshot.result : {
            sanitary: { total: 0, recommendedSize: '', warning: '', codeBasis: 'IPC', drainageOrientation: 'HORIZONTAL' },
            water: { total: 0, flushType: 'FLUSH_TANK', designLengthFt: APP_DEFAULTS.wsfuDesignLengthFt || 100, selectedFlowGpm: 0, serviceSize: '', warning: '', source: '' },
            fixtureRows: [],
            timestamp: ''
          }
        },
        duct: {
          zones: getDuctZonesState(),
          selection: getDuctSelectionState(),
          result: (latestDuctSnapshot && latestDuctSnapshot.result) ? latestDuctSnapshot.result : {
            status: '',
            message: '',
            equivalentDiameter: 0,
            recommendedRound: '',
            recommendedRect: '',
            velocityFpm: 0,
            velocityPressure: 0,
            frictionRate: 0,
            warning: ''
          }
        },
        duct_sizing: {
          zones: getDuctZonesState()
        },
        ductStatic: {
          selection: getDuctStaticSelectionState(),
          result: (latestDuctStaticSnapshot && latestDuctStaticSnapshot.result) ? latestDuctStaticSnapshot.result : {
            segments: getDuctStaticSelectionState().segments,
            totalPressureDrop: 0,
            warnings: []
          }
        },
        solar: {
          selection: getSolarSelectionState(),
          result: (latestSolarSnapshot && latestSolarSnapshot.result) ? latestSolarSnapshot.result : {
            status: 'pending_workbook',
            message: 'Enter worksheet inputs to see live Solar results (or press CALCULATE).',
            notes: document.getElementById('solarNotes') ? document.getElementById('solarNotes').value : ''
          }
        },
        refrigerant: {
          selection: getRefrigerantState(),
          result: (latestRefrigerantSnapshot && latestRefrigerantSnapshot.result) ? latestRefrigerantSnapshot.result : {
            status: 'MANUAL REVIEW REQUIRED',
            message: 'Enter classification and charge data to evaluate compliance.',
            timestamp: ''
          }
        }
      };
    }

    function applyProjectPayload(payload) {
      var settings, areas, container, blocks, i, j, area, block;
      var codeBasis, drainType, slopeSelect, slopeFound;
      var subContainer, wallContainer, subRows, wallRows;
      var condensate, vent, ventilation, gas, wsfu, fixtureUnit, duct, ductStatic, solar, refrigerant, projectContext, gasCodeBasis, gasFuelType, activeModuleFromPayload;

      if (!payload || parseInt(payload.schema_version, 10) !== APP_SCHEMA_VERSION) {
        throw new Error('Unsupported project schema version. Expected ' + APP_SCHEMA_VERSION + '.');
      }

      settings = payload.settings || {};
      areas = payload.areas || [];
      condensate = payload.condensate || {};
      vent = payload.vent || {};
      ventilation = payload.ventilation || {};
      gas = payload.gas || {};
      wsfu = payload.wsfu || {};
      fixtureUnit = payload.fixtureUnit || {};
      duct = payload.duct || {};
      ductStatic = payload.ductStatic || {};
      solar = payload.solar || {};
      refrigerant = payload.refrigerant || {};
      projectContext = (payload.project && typeof payload.project === 'object') ? payload.project : ((payload.projectContext && typeof payload.projectContext === 'object') ? payload.projectContext : {});
      savedProjects = [];
      if (projectContext.savedProjects && projectContext.savedProjects.length) {
        for (i = 0; i < projectContext.savedProjects.length; i++) {
          savedProjects.push(normalizeProjectInfo(projectContext.savedProjects[i]));
        }
      }
      activeProject = projectContext.activeProject ? normalizeProjectInfo(projectContext.activeProject) : null;
      if (hasActiveProject()) upsertSavedProject(activeProject);
      if (window.updateAppState) {
        window.updateAppState('activeProject', activeProject);
        window.updateAppState('savedProjects', savedProjects.slice());
      }
      renderCurrentProjectCard();
      if (!areas || areas.length <= 0) {
        areas = [{ mainArea: 0, subAreas: [], sideWalls: [] }];
      }

      codeBasis = settings.codeBasis || 'IPC';
      if (codeBasis !== 'IPC' && codeBasis !== 'CPC') codeBasis = 'IPC';
      drainType = settings.drainType || 'PRIMARY';
      if (drainType !== 'PRIMARY' && drainType !== 'PRIMARY_SECONDARY' && drainType !== 'COMBINED') drainType = 'PRIMARY';

      document.getElementById('codeBasis').value = codeBasis;
      updateSlopeOptions();
      slopeSelect = document.getElementById('slope');
      slopeFound = false;
      if (slopeSelect && settings.slope) {
        for (i = 0; i < slopeSelect.options.length; i++) {
          if (slopeSelect.options[i].value == settings.slope) {
            slopeSelect.selectedIndex = i;
            slopeFound = true;
            break;
          }
        }
      }
      if (!slopeFound && slopeSelect.options.length > 0) slopeSelect.selectedIndex = 0;

      document.getElementById('drainType').value = drainType;
      document.getElementById('zipCode').value = settings.zipCode || '';
      document.getElementById('rainfallRate').value = parseNonNegativeNumber(settings.rainfallRate, 1).toFixed(2);
      setZipLookupStatus('', '');
      latestGasSnapshot = null;
      gasTableCatalog = [];
      gasDrafts = [];
      if (gas && gas.tables && gas.tables.length) {
        for (i = 0; i < gas.tables.length; i++) {
          gasTableCatalog.push(gas.tables[i]);
        }
      } else {
        seedLegacyGasTables();
      }

      var gasSelection = (gas && gas.selection) ? gas.selection : {};
      gasCodeBasis = gasSelection.codeBasis || 'IPC';
      if (gasCodeBasis !== 'IPC' && gasCodeBasis !== 'CPC') gasCodeBasis = 'IPC';
      gasFuelType = gasSelection.gasType || gasSelection.fuelType || 'NATURAL_GAS';
      if (gasFuelType !== 'NATURAL_GAS' && gasFuelType !== 'PROPANE') gasFuelType = 'NATURAL_GAS';
      if (document.getElementById('gasCodeBasis')) document.getElementById('gasCodeBasis').value = gasCodeBasis;
      if (document.getElementById('gasGasType')) document.getElementById('gasGasType').value = gasFuelType;
      if (document.getElementById('gasMaterial')) document.getElementById('gasMaterial').value = gasSelection.material || 'SCHEDULE_40_METALLIC';
      refreshGasCriteriaDropdowns();
      if (document.getElementById('gasInletPressure')) {
        if (gasSelection.inletPressureOption === 'LT2') document.getElementById('gasInletPressure').value = 'LT2';
        else document.getElementById('gasInletPressure').value = formatCriteriaValue(parseNonNegativeNumber(gasSelection.inletPressure, 0));
      }
      if (document.getElementById('gasPressureDrop')) document.getElementById('gasPressureDrop').value = formatCriteriaValue(parseNonNegativeNumber(gasSelection.pressureDrop, 0));
      if (document.getElementById('gasSpecificGravity')) document.getElementById('gasSpecificGravity').value = formatCriteriaValue(parseNonNegativeNumber(gasSelection.specificGravity, getDefaultSpecificGravityForFuel(gasFuelType)));
      refreshGasCriteriaDropdowns();
      if (document.getElementById('gasDemandMode')) document.getElementById('gasDemandMode').value = (gasSelection.demandMode === 'BTUH' ? 'BTUH' : 'CFH');
      if (document.getElementById('gasHeatingValue')) document.getElementById('gasHeatingValue').value = parseNonNegativeNumber(gasSelection.heatingValue, 0);
      clearGasLineRows();
      var gasLines = (gas && gas.lines && gas.lines.length) ? gas.lines : [{ id: makeGasLineId(), label: 'Line 1', demandValue: 0, runLength: 0 }];
      for (i = 0; i < gasLines.length; i++) addGasLineRow(gasLines[i]);

      if (gas.result && typeof gas.result === 'object') {
        latestGasSnapshot = gas.result;
      }
      if (document.getElementById('gasResults')) document.getElementById('gasResults').innerHTML = 'Enter gas inputs and press Calculate.';

      var condSelection = (condensate && condensate.selection) ? condensate.selection : {};
      var condResult = (condensate && condensate.result) ? condensate.result : null;
      var condZones = (condSelection && condSelection.zones && condSelection.zones.length) ? condSelection.zones : null;
      if (document.getElementById('condSizingMode')) document.getElementById('condSizingMode').value = 'PER_UNIT_QTY';
      condensateZones = [];
      condensateZoneIdSeq = 0;
      condensateRowIdSeq = 0;
      if (!condZones || condZones.length <= 0) {
        var legacyRows = (condSelection && condSelection.equipmentRows && condSelection.equipmentRows.length) ? condSelection.equipmentRows : [];
        if (legacyRows.length <= 0) {
          legacyRows = [{
            id: makeCondensateRowId(),
            equipment: (condSelection.equipment || 'AHU'),
            tonsPerUnit: parseNonNegativeNumber(condSelection.tonsPerUnit, 0),
            quantity: parseNonNegativeNumber(condSelection.quantity, 1)
          }];
        }
        condZones = [{ id: makeCondensateZoneId(), name: 'Zone 1', equipmentRows: legacyRows }];
      }
      for (i = 0; i < condZones.length; i++) {
        addCondensateZone({
          id: condZones[i].id || makeCondensateZoneId(),
          name: condZones[i].name || ('Zone ' + (i + 1)),
          equipmentRows: (condZones[i].equipmentRows && condZones[i].equipmentRows.length) ? condZones[i].equipmentRows : [{ id: makeCondensateRowId(), equipment: 'AHU', tonsPerUnit: 0, quantity: (APP_DEFAULTS.condensateQty || 1) }]
        });
      }
      condensateZoneIdSeq = condensateZones.length;
      condensateRowIdSeq = condensateZones.reduce(function (acc, z) { return acc + ((z && z.equipmentRows) ? z.equipmentRows.length : 0); }, 0);
      updateCondensatePreview();
      if (condResult && typeof condResult === 'object') {
        latestCondensateSnapshot = condResult;
        if (document.getElementById('condResults')) {
          if (!latestCondensateSnapshot.zones || latestCondensateSnapshot.zones.length <= 0) {
            if (latestCondensateSnapshot.rows && latestCondensateSnapshot.rows.length > 0) {
              latestCondensateSnapshot.zones = [{
                id: 'zone_1',
                name: 'Zone 1',
                rows: latestCondensateSnapshot.rows,
                zoneTotalTons: parseNonNegativeNumber(latestCondensateSnapshot.totalTonsUsed, 0),
                recommendedSize: latestCondensateSnapshot.recommendedSize || '',
                warning: latestCondensateSnapshot.warning || ''
              }];
            } else {
              var liveCalc = computeCondensateSnapshot();
              latestCondensateSnapshot.zones = liveCalc.ok ? liveCalc.zones : [];
            }
          }
          document.getElementById('condResults').innerHTML = renderCondensateResultHtml(latestCondensateSnapshot);
        }
      } else {
        latestCondensateSnapshot = null;
        if (document.getElementById('condResults')) document.getElementById('condResults').innerHTML = '<span class="cond-results-placeholder">Condensate flow and pipe sizing results will appear here after calculation.</span>';
      }

      var ventSelection = (vent && vent.selection) ? vent.selection : {};
      var ventResult = (vent && vent.result) ? vent.result : null;
      if (document.getElementById('ventCodeBasis')) document.getElementById('ventCodeBasis').value = ((ventSelection.codeBasis === 'CPC') ? 'CPC' : 'IPC');
      if (document.getElementById('ventUsageType')) document.getElementById('ventUsageType').value = (ventSelection.usageType === 'PUBLIC' ? 'PUBLIC' : 'PRIVATE');
      if (document.getElementById('ventDrainageOrientation')) {
        var loadedOrientation = (ventSelection.drainageOrientation === 'VERTICAL' ? 'VERTICAL' : 'HORIZONTAL');
        if (ventResult && ventResult.drainageOrientation === 'VERTICAL') loadedOrientation = 'VERTICAL';
        document.getElementById('ventDrainageOrientation').value = loadedOrientation;
      }
      if (document.getElementById('ventUseManualTotal')) document.getElementById('ventUseManualTotal').checked = !!ventSelection.useManualTotal;
      if (document.getElementById('ventManualTotalDfu')) document.getElementById('ventManualTotalDfu').value = parseNonNegativeNumber(ventSelection.manualTotalDfu, 0);
      ventFixtureRows = [];
      ventFixtureIdSeq = 0;
      sanitaryZones = [];
      sanitaryZoneIdSeq = 0;
      var loadedSanitaryZones = (ventSelection.zones && ventSelection.zones.length) ? ventSelection.zones : ((vent.zones && vent.zones.length) ? vent.zones : []);
      if (loadedSanitaryZones.length <= 0 && payload.sanitary_drainage && payload.sanitary_drainage.zones && payload.sanitary_drainage.zones.length) {
        loadedSanitaryZones = payload.sanitary_drainage.zones;
      }
      if (loadedSanitaryZones.length <= 0) {
        var loadedVentRows = (ventSelection.fixtureRows && ventSelection.fixtureRows.length) ? ventSelection.fixtureRows : [];
        loadedSanitaryZones = [{ id: makeSanitaryZoneId(), name: 'Zone 1', fixtures: loadedVentRows }];
      }
      for (i = 0; i < loadedSanitaryZones.length; i++) {
        addSanitaryZone(loadedSanitaryZones[i]);
      }
      updateVentTotalsAndUI();
      if (ventResult && typeof ventResult === 'object') {
        latestVentSnapshot = ventResult;
        if (document.getElementById('ventResults')) {
          document.getElementById('ventResults').innerHTML = renderVentResultHtml(ventResult);
        }
      } else {
        latestVentSnapshot = null;
        if (document.getElementById('ventResults')) document.getElementById('ventResults').innerHTML = 'Enter fixture rows and press Calculate.';
      }

      if (ventilation && ventilation.selection && typeof ventilation.selection === 'object') {
        applyVentilationSelectionState(ventilation.selection);
      } else {
        clearVentilationSavedZones();
      }
      if (ventilation && ventilation.result && typeof ventilation.result === 'object') {
        latestVentilationSnapshot = { section: 'ventilation', selection: getVentilationSelectionState(), result: ventilation.result };
        if (document.getElementById('ventilationResults')) {
          document.getElementById('ventilationResults').innerHTML = renderVentilationResultHtml(ventilation.result);
        }
      } else {
        latestVentilationSnapshot = null;
      }
      renderVentilationZonesTable();
      calculateVentilation(false);

      if (document.getElementById('wsfuTotalFu')) {
        document.getElementById('wsfuTotalFu').value = parseNonNegativeNumber((wsfu.selection || {}).totalFu, 0);
      }
      if (document.getElementById('wsfuFlushType')) {
        document.getElementById('wsfuFlushType').value = (((wsfu.selection || {}).flushType === 'FLUSH_VALVE') ? 'FLUSH_VALVE' : 'FLUSH_TANK');
      }
      if (document.getElementById('wsfuUsageType')) {
        document.getElementById('wsfuUsageType').value = (((wsfu.selection || {}).usageType === 'ASSEMBLY') ? 'ASSEMBLY' : (((wsfu.selection || {}).usageType === 'PUBLIC') ? 'PUBLIC' : 'PRIVATE'));
      }
      if (document.getElementById('wsfuDesignLength')) {
        document.getElementById('wsfuDesignLength').value = parseNonNegativeNumber((wsfu.selection || {}).designLengthFt, APP_DEFAULTS.wsfuDesignLengthFt || 100);
      }
      if (document.getElementById('wsfuUseManualTotal')) {
        document.getElementById('wsfuUseManualTotal').checked = !!((wsfu.selection || {}).useManualTotal);
      }
      if (document.getElementById('wsfuManualTotalFu')) {
        document.getElementById('wsfuManualTotalFu').value = parseNonNegativeNumber((wsfu.selection || {}).manualTotalFu, 0);
      }
      wsfuFixtureRows = [];
      wsfuFixtureIdSeq = 1;
      wsfuZones = [];
      wsfuZoneIdSeq = 0;
      var wsfuRows = ((wsfu.selection || {}).fixtureRows && (wsfu.selection || {}).fixtureRows.length) ? (wsfu.selection || {}).fixtureRows : [];
      var loadedWsfuZones = ((wsfu.selection || {}).zones && (wsfu.selection || {}).zones.length) ? (wsfu.selection || {}).zones : ((wsfu.zones && wsfu.zones.length) ? wsfu.zones : []);
      var legacyTotalFu = parseNonNegativeNumber((wsfu.selection || {}).totalFu, 0);
      if (wsfuRows.length <= 0 && legacyTotalFu > 0) {
        if (document.getElementById('wsfuUseManualTotal')) document.getElementById('wsfuUseManualTotal').checked = true;
        if (document.getElementById('wsfuManualTotalFu')) document.getElementById('wsfuManualTotalFu').value = legacyTotalFu;
        wsfuRows = [{ id: makeWsfuFixtureId(), fixtureKey: (wsfuFixtureCatalog.length > 0 ? wsfuFixtureCatalog[0].key : ''), quantity: 1, hotCold: false, unitWsfu: 0, rowTotalWsfu: 0 }];
      }
      if (loadedWsfuZones.length <= 0) loadedWsfuZones = [{ id: makeWsfuZoneId(), name: 'Zone 1', fixtures: wsfuRows }];
      for (i = 0; i < loadedWsfuZones.length; i++) addWsfuZone(loadedWsfuZones[i]);
      updateWsfuTotalsAndUI();
      latestWsfuSnapshot = (wsfu.result && typeof wsfu.result === 'object') ? { section: 'wsfu', selection: (wsfu.selection || {}), result: wsfu.result } : null;
      if (document.getElementById('wsfuResults')) {
        document.getElementById('wsfuResults').innerHTML = latestWsfuSnapshot ? renderWsfuResult(latestWsfuSnapshot.result) : 'Enter fixture rows and press Calculate.';
      }

      var fixtureSelection = (fixtureUnit && fixtureUnit.selection) ? fixtureUnit.selection : {};
      if (document.getElementById('fixtureUnitCodeBasis')) {
        document.getElementById('fixtureUnitCodeBasis').value = ((fixtureSelection.codeBasis === 'CPC') ? 'CPC' : 'IPC');
      }
      if (document.getElementById('fixtureUnitDrainageOrientation')) {
        document.getElementById('fixtureUnitDrainageOrientation').value = (fixtureSelection.drainageOrientation === 'VERTICAL' ? 'VERTICAL' : 'HORIZONTAL');
      }
      if (document.getElementById('fixtureUnitFlushType')) {
        document.getElementById('fixtureUnitFlushType').value = (fixtureSelection.flushType === 'FLUSH_VALVE' ? 'FLUSH_VALVE' : 'FLUSH_TANK');
      }
      if (document.getElementById('fixtureUnitDesignLength')) {
        document.getElementById('fixtureUnitDesignLength').value = parseNonNegativeNumber(fixtureSelection.designLengthFt, APP_DEFAULTS.wsfuDesignLengthFt || 100);
      }
      fixtureUnitZones = [];
      fixtureUnitZoneIdSeq = 0;
      fixtureUnitRowIdSeq = 0;
      var loadedFixtureZones = (fixtureSelection.zones && fixtureSelection.zones.length) ? fixtureSelection.zones : [];
      if (loadedFixtureZones.length <= 0) {
        var legacyRows = (fixtureSelection.rows && fixtureSelection.rows.length) ? fixtureSelection.rows : [];
        loadedFixtureZones = [{ id: makeFixtureUnitZoneId(), name: 'Zone 1', rows: legacyRows }];
      }
      for (i = 0; i < loadedFixtureZones.length; i++) addFixtureUnitZone(loadedFixtureZones[i]);
      if (fixtureUnitZones.length <= 0) fixtureUnitZones = [createDefaultFixtureUnitZone(1)];
      latestFixtureUnitSnapshot = (fixtureUnit.result && typeof fixtureUnit.result === 'object') ? { section: 'fixtureUnit', selection: fixtureSelection, result: fixtureUnit.result } : null;
      if (latestFixtureUnitSnapshot && latestFixtureUnitSnapshot.result) renderFixtureUnitResults(latestFixtureUnitSnapshot.result);
      else clearFixtureUnitResults();
      updateFixtureUnitRowsAndPreview();

      var ductSelection = (duct && duct.selection) ? duct.selection : {};
      ductZones = [];
      ductZoneIdSeq = 0;
      ductSegmentIdSeq = 0;
      var loadedDuctZones = (ductSelection.zones && ductSelection.zones.length) ? ductSelection.zones : ((duct.zones && duct.zones.length) ? duct.zones : []);
      if (loadedDuctZones.length <= 0 && payload.duct_sizing && payload.duct_sizing.zones && payload.duct_sizing.zones.length) {
        loadedDuctZones = payload.duct_sizing.zones;
      }
      if (loadedDuctZones.length <= 0) {
        loadedDuctZones = [{
          id: makeDuctZoneId(),
          name: 'Zone 1',
          segments: [{
            id: makeDuctSegmentId(),
            mode: normalizeDuctMode((ductSelection.mode || DUCT_CALC.modes.SIZE_FROM_CFM_FRICTION), (ductSelection.shape || DUCT_CALC.shapes.ROUND)),
            shape: (ductSelection.shape || DUCT_CALC.shapes.ROUND),
            cfm: parseNonNegativeNumber(ductSelection.cfm, 0),
            frictionTarget: parseNonNegativeNumber(ductSelection.frictionTarget, 0),
            velocityTarget: parseNonNegativeNumber(ductSelection.velocityTarget, 0),
            diameter: parseNonNegativeNumber(ductSelection.diameter, 0),
            width: parseNonNegativeNumber(ductSelection.width, 0),
            height: parseNonNegativeNumber(ductSelection.height, 0),
            ratioLimit: parseNonNegativeNumber(ductSelection.ratioLimit, 3)
          }]
        }];
      }
      for (i = 0; i < loadedDuctZones.length; i++) addDuctZone(loadedDuctZones[i]);
      renderDuctZones();
      latestDuctSnapshot = (duct.result && typeof duct.result === 'object') ? { section: 'duct', selection: ductSelection, result: duct.result } : null;
      if (document.getElementById('ductResults')) {
        document.getElementById('ductResults').innerHTML = latestDuctSnapshot ? renderDuctResult(latestDuctSnapshot.result) : '<div class="result-placeholder">' + (APP_MESSAGES.ductPlaceholder || 'Enter airflow and sizing criteria to see live duct size recommendations.') + '</div>';
      }

      ductStaticSegments = [];
      ductStaticIdSeq = 1;
      var ductStaticSelection = (ductStatic && ductStatic.selection) ? ductStatic.selection : {};
      var ductStaticResult = (ductStatic && ductStatic.result) ? ductStatic.result : null;
      var restoredSegments = (ductStaticSelection && ductStaticSelection.segments && ductStaticSelection.segments.length) ? ductStaticSelection.segments : [];
      if (restoredSegments.length <= 0 && ductStaticResult && ductStaticResult.segments && ductStaticResult.segments.length) {
        restoredSegments = ductStaticResult.segments;
      }
      if (restoredSegments.length <= 0) {
        addDuctStaticSegment();
      } else {
        for (i = 0; i < restoredSegments.length; i++) {
          addDuctStaticSegment(restoredSegments[i]);
        }
      }
      if (ductStaticResult && typeof ductStaticResult === 'object') {
        latestDuctStaticSnapshot = { section: 'ductstatic', selection: getDuctStaticSelectionState(), result: ductStaticResult };
        if (document.getElementById('ductStaticResults')) {
          document.getElementById('ductStaticResults').innerHTML = renderDuctStaticResult(ductStaticResult);
        }
      } else {
        latestDuctStaticSnapshot = null;
        if (document.getElementById('ductStaticResults')) {
          document.getElementById('ductStaticResults').innerHTML = '<div class="result-placeholder">' + (APP_MESSAGES.ductStaticPlaceholder || 'Enter duct segment data to calculate total static pressure drop.') + '</div>';
        }
      }

      if (document.getElementById('solarBuildingSf')) document.getElementById('solarBuildingSf').value = parseNonNegativeNumber((solar.selection || {}).buildingSf, 0);
      if (document.getElementById('solarZip')) document.getElementById('solarZip').value = (solar.selection && solar.selection.zipCode) ? solar.selection.zipCode : '';
      if (document.getElementById('solarClimateZone')) document.getElementById('solarClimateZone').value = parseNonNegativeNumber((solar.selection || {}).climateZone, 0);
      if (document.getElementById('solarBuildingType') && solar.selection && solar.selection.buildingType) document.getElementById('solarBuildingType').value = solar.selection.buildingType;
      if (document.getElementById('solarUseSaraPath')) document.getElementById('solarUseSaraPath').checked = !!(solar.selection && solar.selection.useSaraPath);
      if (document.getElementById('solarSara')) document.getElementById('solarSara').value = parseNonNegativeNumber((solar.selection || {}).sara, 0);
      if (document.getElementById('solarDValue')) document.getElementById('solarDValue').value = parseNonNegativeNumber((solar.selection || {}).dValue, 0);
      if (document.getElementById('solarExempt')) document.getElementById('solarExempt').checked = !!(solar.selection && solar.selection.solarExempt);
      if (document.getElementById('batteryExempt')) document.getElementById('batteryExempt').checked = !!(solar.selection && solar.selection.batteryExempt);
      if (document.getElementById('solarNotes')) document.getElementById('solarNotes').value = (solar.selection && solar.selection.notes) ? solar.selection.notes : '';
      solarClimateManualOverride = parseNonNegativeNumber((solar.selection || {}).climateZone, 0) > 0;
      onSolarSaraToggle();
      latestSolarSnapshot = (solar.result && typeof solar.result === 'object') ? { section: 'solar', selection: getSolarSelectionState(), result: solar.result } : null;
      if (document.getElementById('solarResults')) {
        if (latestSolarSnapshot) {
          document.getElementById('solarResults').innerHTML = renderSolarResult(latestSolarSnapshot.result);
        } else {
          document.getElementById('solarResults').innerHTML = 'Enter worksheet inputs to see live Solar results (or press CALCULATE).';
        }
      }

      if (refrigerant && refrigerant.selection && typeof refrigerant.selection === 'object') {
        var refSel = refrigerant.selection;
        if (document.getElementById('refCalculationMode')) document.getElementById('refCalculationMode').value = refSel.calculationMode || 'SINGLE';
        if (document.getElementById('refRefrigerantType')) document.getElementById('refRefrigerantType').value = refSel.classification ? (refSel.classification.refrigerantType || 'R-454B') : 'R-454B';
        if (document.getElementById('refSystemType')) document.getElementById('refSystemType').value = refSel.classification ? (refSel.classification.systemType || 'SPLIT_DX') : 'SPLIT_DX';
        if (document.getElementById('refSpaceType')) document.getElementById('refSpaceType').value = refSel.classification ? (refSel.classification.spaceType || 'RESIDENTIAL') : 'RESIDENTIAL';
        if (document.getElementById('refOneSystemOneDwelling')) document.getElementById('refOneSystemOneDwelling').value = refSel.classification ? (refSel.classification.oneSystemOneDwelling || 'YES') : 'YES';
        if (document.getElementById('refHighProbability')) document.getElementById('refHighProbability').value = refSel.classification ? (refSel.classification.highProbability || 'YES') : 'YES';
        if (document.getElementById('refDataVersion')) document.getElementById('refDataVersion').value = refSel.classification ? (refSel.classification.dataVersion || 'ADDENDA') : 'ADDENDA';

        if (refSel.chargeInputs) {
          if (document.getElementById('refFactoryCharge')) document.getElementById('refFactoryCharge').value = parseNonNegativeNumber(refSel.chargeInputs.factoryCharge, 0);
          if (document.getElementById('refFieldCharge')) document.getElementById('refFieldCharge').value = parseNonNegativeNumber(refSel.chargeInputs.fieldCharge, 0);
          if (document.getElementById('refInstalledPipeLength')) document.getElementById('refInstalledPipeLength').value = parseNonNegativeNumber(refSel.chargeInputs.installedPipeLength, 0);
          if (document.getElementById('refIncludedPipeLength')) document.getElementById('refIncludedPipeLength').value = parseNonNegativeNumber(refSel.chargeInputs.includedPipeLength, 0);
          if (document.getElementById('refAdditionalChargeRate')) document.getElementById('refAdditionalChargeRate').value = parseNonNegativeNumber(refSel.chargeInputs.additionalChargeRate, 0);
          if (document.getElementById('refUseManualAdditionalCharge')) document.getElementById('refUseManualAdditionalCharge').checked = !!refSel.chargeInputs.useManualAdditionalCharge;
          if (document.getElementById('refManualAdditionalCharge')) document.getElementById('refManualAdditionalCharge').value = parseNonNegativeNumber(refSel.chargeInputs.manualAdditionalCharge, 0);
        }

        if (refSel.ashrae15Inputs) {
          if (document.getElementById('refEqSelectionMode')) document.getElementById('refEqSelectionMode').value = refSel.ashrae15Inputs.selectionMode || 'AUTO';
          if (document.getElementById('refConnectedSpaces')) document.getElementById('refConnectedSpaces').value = refSel.ashrae15Inputs.connectedSpaces || 'NO';
          if (document.getElementById('refAirCirculation')) document.getElementById('refAirCirculation').value = refSel.ashrae15Inputs.airCirculation || 'YES';
          if (document.getElementById('refManualEquation')) document.getElementById('refManualEquation').value = refSel.ashrae15Inputs.manualEquation || 'EQ_7_8';
          if (document.getElementById('refRoomArea78')) document.getElementById('refRoomArea78').value = parseNonNegativeNumber(refSel.ashrae15Inputs.roomArea78, 0);
          if (document.getElementById('refCeilingHeight78')) document.getElementById('refCeilingHeight78').value = parseNonNegativeNumber(refSel.ashrae15Inputs.ceilingHeight78, 0);
          if (document.getElementById('refLfl78')) document.getElementById('refLfl78').value = parseNonNegativeNumber(refSel.ashrae15Inputs.lfl78, 0);
          if (document.getElementById('refCf78')) document.getElementById('refCf78').value = parseNonNegativeNumber(refSel.ashrae15Inputs.cf78, 0.5);
          if (document.getElementById('refFocc78')) document.getElementById('refFocc78').value = parseNonNegativeNumber(refSel.ashrae15Inputs.focc78, 1);
          if (document.getElementById('refMdef79')) document.getElementById('refMdef79').value = parseNonNegativeNumber(refSel.ashrae15Inputs.mdef79, 0);
          if (document.getElementById('refFlfl79')) document.getElementById('refFlfl79').value = parseNonNegativeNumber(refSel.ashrae15Inputs.flfl79, 1);
          if (document.getElementById('refFocc79')) document.getElementById('refFocc79').value = parseNonNegativeNumber(refSel.ashrae15Inputs.focc79, 1);
          if (document.getElementById('refLowestOpeningHeight79')) document.getElementById('refLowestOpeningHeight79').value = parseNonNegativeNumber(refSel.ashrae15Inputs.lowestOpeningHeight79, 0);
          if (document.getElementById('refOptionalRoomArea79')) document.getElementById('refOptionalRoomArea79').value = parseNonNegativeNumber(refSel.ashrae15Inputs.optionalRoomArea79, 0);
        }

        if (refSel.ashrae152Inputs) {
          if (document.getElementById('refM1_152')) document.getElementById('refM1_152').value = parseNonNegativeNumber(refSel.ashrae152Inputs.m1, 0);
          if (document.getElementById('refM2_152')) document.getElementById('refM2_152').value = parseNonNegativeNumber(refSel.ashrae152Inputs.m2, 0);
          if (document.getElementById('refRoomArea152')) document.getElementById('refRoomArea152').value = parseNonNegativeNumber(refSel.ashrae152Inputs.roomArea, 0);
          if (document.getElementById('refRoomVolume152')) document.getElementById('refRoomVolume152').value = parseNonNegativeNumber(refSel.ashrae152Inputs.roomVolume, 0);
          if (document.getElementById('refSpaceHeight152')) document.getElementById('refSpaceHeight152').value = parseNonNegativeNumber(refSel.ashrae152Inputs.spaceHeight || refSel.ashrae152Inputs.dispersalHeight, 7.2);
          if (document.getElementById('refVentilationCfm152')) document.getElementById('refVentilationCfm152').value = parseNonNegativeNumber(refSel.ashrae152Inputs.ventilationCfm, 0);
          if (document.getElementById('refInterpolate152')) document.getElementById('refInterpolate152').value = refSel.ashrae152Inputs.interpolateLookups === false ? 'NO' : 'YES';
          if (document.getElementById('refAllowAreaExtrapolation152')) document.getElementById('refAllowAreaExtrapolation152').value = refSel.ashrae152Inputs.allowAreaExtrapolation ? 'YES' : 'NO';
          if (document.getElementById('refUseCOverride152')) document.getElementById('refUseCOverride152').value = refSel.ashrae152Inputs.useCOverride ? 'YES' : 'NO';
          if (document.getElementById('refCOverride152')) document.getElementById('refCOverride152').value = parseNonNegativeNumber(refSel.ashrae152Inputs.cOverride, 0);
          if (document.getElementById('refVentilation152')) document.getElementById('refVentilation152').value = refSel.ashrae152Inputs.ventilationUsed || 'NO';
          if (document.getElementById('refSsov152')) document.getElementById('refSsov152').value = refSel.ashrae152Inputs.ssovUsed || 'NO';
        }

        if (refSel.advancedInputs) {
          if (document.getElementById('refUseMrel152')) document.getElementById('refUseMrel152').value = refSel.advancedInputs.useMrel || 'NO';
          if (document.getElementById('refMrel152')) document.getElementById('refMrel152').value = parseNonNegativeNumber(refSel.advancedInputs.mrel, 0);
        }

        if (refSel.multiZone && refSel.multiZone.shared) {
          var refShared = refSel.multiZone.shared;
          if (document.getElementById('refMzRefrigerantType')) document.getElementById('refMzRefrigerantType').value = refShared.refrigerantType || 'R-454B';
          if (document.getElementById('refMzSafetyGroup')) document.getElementById('refMzSafetyGroup').value = refShared.safetyGroup || 'A2L';
          if (document.getElementById('refMzStandardPath')) document.getElementById('refMzStandardPath').value = refShared.standardPath || 'AUTO';
          if (document.getElementById('refMzDataVersion')) document.getElementById('refMzDataVersion').value = refShared.dataVersion || 'ADDENDA';
          if (document.getElementById('refMzDefaultPipeSize')) document.getElementById('refMzDefaultPipeSize').value = refShared.defaultPipeSize || '3/8 in';
          if (document.getElementById('refMzDefaultChargePerFt')) document.getElementById('refMzDefaultChargePerFt').value = parseNonNegativeNumber(refShared.defaultChargePerFt, 0);
          if (document.getElementById('refMzDefaultFactoryCharge')) document.getElementById('refMzDefaultFactoryCharge').value = parseNonNegativeNumber(refShared.defaultFactoryCharge, 0);
          if (document.getElementById('refMzVentilationConsidered')) document.getElementById('refMzVentilationConsidered').value = refShared.ventilationConsidered || 'NO';
          if (document.getElementById('refMzSsovConsidered')) document.getElementById('refMzSsovConsidered').value = refShared.ssovConsidered || 'NO';
          if (document.getElementById('refMzDefaultUnits')) document.getElementById('refMzDefaultUnits').value = parseNonNegativeNumber(refShared.defaultUnits, 1);
          if (document.getElementById('refMzDefaultSpaceHeight')) document.getElementById('refMzDefaultSpaceHeight').value = parseNonNegativeNumber(refShared.defaultSpaceHeight, 7.2);
          if (document.getElementById('refMzDefaultVentCfm')) document.getElementById('refMzDefaultVentCfm').value = parseNonNegativeNumber(refShared.defaultVentCfm, 0);
          if (document.getElementById('refMzDefaultM1')) document.getElementById('refMzDefaultM1').value = parseNonNegativeNumber(refShared.defaultM1, 0);
          if (document.getElementById('refMzDefaultM2')) document.getElementById('refMzDefaultM2').value = parseNonNegativeNumber(refShared.defaultM2, 0);
          if (document.getElementById('refMzInterpolate')) document.getElementById('refMzInterpolate').value = refShared.interpolateLookups === false ? 'NO' : 'YES';
          if (document.getElementById('refMzAllowAreaExtrapolation')) document.getElementById('refMzAllowAreaExtrapolation').value = refShared.allowAreaExtrapolation ? 'YES' : 'NO';
        }
        refrigerantMultiRows = [];
        refrigerantMultiRowIdSeq = 0;
        if (refSel.multiZone && refSel.multiZone.rows && refSel.multiZone.rows.length) {
          var mzRows = refSel.multiZone.rows;
          for (i = 0; i < mzRows.length; i++) addRefrigerantMultiRow(mzRows[i]);
        } else {
          addRefrigerantMultiRow();
        }
      }
      if (refrigerant && refrigerant.result && typeof refrigerant.result === 'object') {
        latestRefrigerantSnapshot = { section: 'refrigerant', selection: getRefrigerantState(), result: refrigerant.result };
      } else {
        latestRefrigerantSnapshot = null;
      }
      calculateRefrigerantCompliance(false);

      container = document.getElementById('areasContainer');
      container.innerHTML = '';
      for (i = 0; i < areas.length; i++) {
        area = areas[i] || {};
        container.insertAdjacentHTML('beforeend', makeAreaBlock((area && area.name) ? String(area.name) : getAreaLabel(i), parseNonNegativeNumber(area.mainArea, 0)));
      }

      blocks = getAreaBlocks();
      for (i = 0; i < blocks.length; i++) {
        area = areas[i] || {};
        block = blocks[i];

        var areaInput = block.getElementsByClassName('area-main-input');
        if (areaInput.length > 0) areaInput[0].value = parseNonNegativeNumber(area.mainArea, 0);

        subContainer = block.getElementsByClassName('subareas-container');
        if (subContainer.length > 0 && area.subAreas && area.subAreas.length > 0) {
          for (j = 0; j < area.subAreas.length; j++) {
            subContainer[0].insertAdjacentHTML('beforeend', makeSubAreaRow());
            subRows = getSubAreaRows(block);
            if (subRows.length > 0) {
              var subInput = subRows[subRows.length - 1].getElementsByClassName('subarea-input');
              if (subInput.length > 0) subInput[0].value = parseNonNegativeNumber(area.subAreas[j], 0);
            }
          }
        }

        wallContainer = block.getElementsByClassName('sidewalls-container');
        if (wallContainer.length > 0 && area.sideWalls && area.sideWalls.length > 0) {
          for (j = 0; j < area.sideWalls.length; j++) {
            wallContainer[0].insertAdjacentHTML('beforeend', makeSideWallRow());
            wallRows = getSideWallRows(block);
            if (wallRows.length > 0) {
              var wallInput = wallRows[wallRows.length - 1].getElementsByClassName('sidewall-input');
              if (wallInput.length > 0) wallInput[0].value = parseNonNegativeNumber(area.sideWalls[j], 0);
            }
          }
        }

        updateSubAreaMetadata(block);
        updateSideWallMetadata(block);
      }

      renumberAreas();
      recalculate(false);
      updateCondensatePreview();
      updateVentTotalsAndUI();
      calculateVentilation(false);
      calculateGasPipe(false);
      updateWsfuTotalsAndUI();
      calculateDuct(false);
      updateSolarPreview();
      activeModuleFromPayload = payload.activeModule || 'home';
      if (activeModuleFromPayload !== 'home' && activeModuleFromPayload !== 'storm' && activeModuleFromPayload !== 'condensate' && activeModuleFromPayload !== 'vent' && activeModuleFromPayload !== 'ventilation' && activeModuleFromPayload !== 'gas' && activeModuleFromPayload !== 'wsfu' && activeModuleFromPayload !== 'fixtureUnit' && activeModuleFromPayload !== 'solar' && activeModuleFromPayload !== 'duct' && activeModuleFromPayload !== 'ductStatic' && activeModuleFromPayload !== 'refrigerant') {
        activeModuleFromPayload = 'home';
      }
      setActiveModule(activeModuleFromPayload);
    }

    async function saveProjectToFile() {
      if (!confirmProjectContextForAction('project save')) return;
      var payload = buildProjectPayload();
      if (!window.pywebview || !window.pywebview.api || !window.pywebview.api.save_project) {
        alert('Project save service is not ready yet. Please wait a moment and try again.');
        return;
      }
      try {
        var result = await window.pywebview.api.save_project(payload);
        if (result && result.ok) {
          alert('Project saved successfully.\\n' + result.path);
          return;
        }
        if (result && result.reason && result.reason.indexOf('canceled') >= 0) return;
        alert('Project save failed.\\n' + (result && result.reason ? result.reason : 'Unknown error'));
      } catch (e) {
        alert('Project save failed.\\n' + (e && e.message ? e.message : 'Unknown error'));
      }
    }

    async function loadProjectFromFile() {
      if (!window.pywebview || !window.pywebview.api || !window.pywebview.api.load_project) {
        alert('Project load service is not ready yet. Please wait a moment and try again.');
        return;
      }
      try {
        var result = await window.pywebview.api.load_project();
        if (!result || !result.ok) {
          if (result && result.reason && result.reason.indexOf('canceled') >= 0) return;
          alert('Project load failed.\\n' + (result && result.reason ? result.reason : 'Unknown error'));
          return;
        }
        applyProjectPayload(result.data);
        alert('Project loaded successfully.\\n' + result.path);
      } catch (e) {
        alert('Project load failed.\\n' + (e && e.message ? e.message : 'Unknown error'));
      }
    }

    async function exportPerAreaToExcel() {
      if (!confirmProjectContextForAction('export')) return;
      var snapshot = latestPerAreaExport ? Object.assign({}, latestPerAreaExport, { project: getProjectContextPayload() }) : null;
      if (!snapshot || !snapshot.rows || snapshot.rows.length <= 0) {
        alert('No valid per-area results are available to export yet. Please calculate first.');
        return;
      }
      if (!window.pywebview || !window.pywebview.api || !window.pywebview.api.export_per_area) {
        alert('Export service is not ready yet. Please wait a moment and try again.');
        return;
      }
      try {
        var result = await window.pywebview.api.export_per_area(snapshot);
        if (result && result.ok) {
          alert('Per-Area Results exported successfully. Rows: ' + result.rowCount + '. Saved to: ' + result.path);
          return;
        }
        alert('Export failed.\\n' + (result && result.reason ? result.reason : 'Unknown error'));
      } catch (e) {
        alert('Export failed.\\n' + (e && e.message ? e.message : 'Unknown error'));
      }
    }

    function buildSectionExportSnapshot() {
      if (activeModule === 'home') {
        var payload = buildProjectPayload();
        return {
          section: 'home',
          project: getProjectContextPayload(),
          summary: {
            activeModule: payload.activeModule || 'home',
            areasCount: (payload.areas || []).length,
            gasLinesCount: (payload.gas && payload.gas.lines) ? payload.gas.lines.length : 0
          }
        };
      }
      if (activeModule === 'condensate') return latestCondensateSnapshot ? Object.assign({}, latestCondensateSnapshot, { project: getProjectContextPayload() }) : null;
      if (activeModule === 'vent') return latestVentSnapshot ? Object.assign({}, latestVentSnapshot, { project: getProjectContextPayload() }) : null;
      if (activeModule === 'ventilation') {
        return {
          section: 'ventilation',
          project: getProjectContextPayload(),
          selection: getVentilationSelectionState(),
          result: (latestVentilationSnapshot && latestVentilationSnapshot.result) ? latestVentilationSnapshot.result : {
            status: 'incomplete',
            message: 'Select occupancy and enter zone data to calculate required outdoor air.',
            zones: []
          }
        };
      }
      if (activeModule === 'gas') {
        if (!latestGasSnapshot) return null;
        return {
          section: 'gas',
          project: getProjectContextPayload(),
          selection: getGasSelectionState(),
          result: latestGasSnapshot
        };
      }
      if (activeModule === 'wsfu') {
        if (latestWsfuSnapshot) {
          return {
            section: 'wsfu',
            project: getProjectContextPayload(),
            selection: buildProjectPayload().wsfu.selection,
            result: latestWsfuSnapshot.result
          };
        }
        return null;
      }
      if (activeModule === 'fixtureUnit') {
        if (latestFixtureUnitSnapshot) {
          return {
            section: 'fixtureunit',
            project: getProjectContextPayload(),
            selection: getFixtureUnitSelectionState(),
            result: latestFixtureUnitSnapshot.result
          };
        }
        return {
          section: 'fixtureunit',
          project: getProjectContextPayload(),
          selection: getFixtureUnitSelectionState(),
          result: {
            sanitary: { total: 0, recommendedSize: '', warning: '' },
            water: { total: 0, selectedFlowGpm: 0, serviceSize: '', warning: '' },
            fixtureRows: []
          }
        };
      }
      if (activeModule === 'duct') {
        if (latestDuctSnapshot) {
          return {
            section: 'duct',
            project: getProjectContextPayload(),
            selection: getDuctSelectionState(),
            result: latestDuctSnapshot.result
          };
        }
        return null;
      }
      if (activeModule === 'ductStatic') {
        if (latestDuctStaticSnapshot) {
          return {
            section: 'ductstatic',
            project: getProjectContextPayload(),
            selection: getDuctStaticSelectionState(),
            result: latestDuctStaticSnapshot.result
          };
        }
        return {
          section: 'ductstatic',
          project: getProjectContextPayload(),
          selection: getDuctStaticSelectionState(),
          result: {
            segments: getDuctStaticSelectionState().segments,
            totalPressureDrop: 0,
            warnings: []
          }
        };
      }
      if (activeModule === 'solar') {
        if (latestSolarSnapshot) {
          return {
            section: 'solar',
            project: getProjectContextPayload(),
            selection: getSolarSelectionState(),
            result: latestSolarSnapshot.result
          };
        }
        return {
          section: 'solar',
          project: getProjectContextPayload(),
          selection: getSolarSelectionState(),
          result: {
            status: 'pending_workbook',
            message: 'Enter worksheet inputs to see live Solar results (or press CALCULATE).',
            notes: document.getElementById('solarNotes') ? document.getElementById('solarNotes').value : ''
          }
        };
      }
      if (activeModule === 'refrigerant') {
        if (latestRefrigerantSnapshot) {
          return {
            section: 'refrigerant',
            project: getProjectContextPayload(),
            selection: getRefrigerantState(),
            result: latestRefrigerantSnapshot.result
          };
        }
        return {
          section: 'refrigerant',
          project: getProjectContextPayload(),
          selection: getRefrigerantState(),
          result: {
            status: 'MANUAL REVIEW REQUIRED',
            message: 'Enter classification and charge data to evaluate compliance.',
            timestamp: ''
          }
        };
      }
      return null;
    }

    async function exportSectionToExcel() {
      if (!confirmProjectContextForAction('export')) return;
      if (activeModule === 'storm') {
        await exportPerAreaToExcel();
        return;
      }
      var snapshot = buildSectionExportSnapshot();
      if (!snapshot) {
        alert('No valid results are available to export in this section yet. Please calculate first.');
        return;
      }
      if (!window.pywebview || !window.pywebview.api || !window.pywebview.api.export_section) {
        alert('Export service is not ready yet. Please wait a moment and try again.');
        return;
      }
      try {
        var result = await window.pywebview.api.export_section(snapshot);
        if (result && result.ok) {
          alert('Section exported successfully. Saved to: ' + result.path);
          return;
        }
        alert('Export failed.\\n' + (result && result.reason ? result.reason : 'Unknown error'));
      } catch (e) {
        alert('Export failed.\\n' + (e && e.message ? e.message : 'Unknown error'));
      }
    }

    function enforceStartupWindowSize() {
      return;
    }

    function setStormPreviewSummary(totalArea, gpmEffective, cfsEffective, areaCount, statusText, isWarning) {
      var totalNode = document.getElementById('stormPreviewArea');
      var runoffNode = document.getElementById('stormPreviewRunoff');
      var countNode = document.getElementById('stormPreviewCount');
      var statusNode = document.getElementById('stormPreviewStatus');
      if (totalNode) totalNode.innerHTML = formatNumber(totalArea, 0) + ' sq ft';
      if (runoffNode) runoffNode.innerHTML = formatNumber(gpmEffective, 2) + ' GPM (' + formatNumber(cfsEffective, 3) + ' CFS)';
      if (countNode) countNode.innerHTML = String(areaCount);
      if (statusNode) {
        statusNode.innerHTML = statusText || '';
        statusNode.className = 'storm-preview-status ' + (isWarning ? 'warn' : 'ok');
      }
    }

    function updateStormResultPlaceholders() {
      var selectedHint = document.getElementById('selectedBasisHint');
      var runoffHint = document.getElementById('runoffHint');
      var perAreaHint = document.getElementById('perAreaHint');
      var selectedHasData = !!String((document.getElementById('selectedCodeText') || {}).innerHTML || '').trim();
      var runoffHasData = !!String((document.getElementById('gpmText') || {}).innerHTML || '').trim();
      var perAreaHasData = !!String((document.getElementById('perAreaResultsList') || {}).innerHTML || '').trim();
      if (selectedHint) selectedHint.style.display = selectedHasData ? 'none' : 'block';
      if (runoffHint) runoffHint.style.display = runoffHasData ? 'none' : 'block';
      if (perAreaHint) perAreaHint.style.display = perAreaHasData ? 'none' : 'block';
    }

    function recalculate(commitResults) {
      if (commitResults !== true) commitResults = false;
      var rainfallInput = document.getElementById('rainfallRate');
      var rainfallRate = parseFloat(document.getElementById('rainfallRate').value);
      var drainType = document.getElementById('drainType').value;
      var drainTypeLabel = document.getElementById('drainType').options[document.getElementById('drainType').selectedIndex].text;
      var demandMultiplier = (drainType == 'COMBINED') ? 2.00 : 1.00;
      var areaData;
      var totalArea;
      var gpmBase;
      var cfsBase;
      var gpmEffective;
      var cfsEffective;
      var areaResultRows = [];
      var i, areaEntry, horizontalResult, verticalResult, warningText, effectiveAreaForSizing;
      var slopeSelect = document.getElementById('slope');
      var slopeLabel = '';
      var subAreaText = '';

      clearValidationSummary();

      if (rainfallInput.value === '' || isNaN(rainfallRate) || rainfallRate <= 0) {
        showInputError(rainfallInput, null, '');
        showValidationSummary('Enter a rainfall rate greater than zero to calculate results.');
        setStormPreviewSummary(0, 0, 0, getAreaBlocks().length, 'Enter rainfall rate > 0 to preview sizing.', true);
        return;
      }

      areaData = parseAreasForRunoff(rainfallRate);
      if (!areaData.isValid) {
        showValidationSummary('Fix highlighted area values before calculating.');
        setStormPreviewSummary(0, 0, 0, getAreaBlocks().length, 'Fix highlighted area values to preview sizing.', true);
        return;
      }
      totalArea = areaData.totalArea;
      var codeBasis = document.getElementById('codeBasis').value;
      var slope = document.getElementById('slope').value;
      var data = pipeData[codeBasis];
      if (slopeSelect && slopeSelect.selectedIndex >= 0) slopeLabel = slopeSelect.options[slopeSelect.selectedIndex].text;
      else slopeLabel = slope;

      gpmBase = rainfallRate * totalArea * 0.0104;
      cfsBase = gpmBase / 448.831;
      gpmEffective = gpmBase * demandMultiplier;
      cfsEffective = gpmEffective / 448.831;

      for (i = 0; i < areaData.perArea.length; i++) {
        areaEntry = areaData.perArea[i];
        effectiveAreaForSizing = areaEntry.area * demandMultiplier;
        horizontalResult = pickPipeSizeWithOverflow(effectiveAreaForSizing, rainfallRate, data.horizontal[slope]);
        verticalResult = null;
        warningText = '';
        if (data.vertical.length > 0) verticalResult = pickPipeSizeWithOverflow(effectiveAreaForSizing, rainfallRate, data.vertical);
        if (horizontalResult.exceeded) {
          warningText += 'Effective area exceeds horizontal table limit (' + formatNumber(horizontalResult.maxAllowableArea, 0) + ' sq ft max). ';
        }
        if (verticalResult !== null && verticalResult.exceeded) {
          warningText += 'Effective area exceeds vertical table limit (' + formatNumber(verticalResult.maxAllowableArea, 0) + ' sq ft max).';
        }
        subAreaText = (areaEntry.subAreaCount > 0) ? areaEntry.subAreaValues.join(', ') : 'None';
        areaResultRows.push({
          label: areaEntry.label,
          area: areaEntry.area,
          roofArea: areaEntry.roofArea,
          sideWallAreas: areaEntry.sideWallAreas,
          sideWallCount: areaEntry.sideWallCount,
          sideWallSum: areaEntry.sideWallSum,
          sideWallFactor: areaEntry.sideWallFactor,
          sideWallContribution: areaEntry.sideWallContribution,
          effectiveDrainageArea: areaEntry.effectiveArea,
          parentArea: areaEntry.parentArea,
          subAreaCount: areaEntry.subAreaCount,
          subAreaValues: areaEntry.subAreaValues,
          subAreaListText: subAreaText,
          baseGpm: areaEntry.gpm,
          baseCfs: areaEntry.gpm / 448.831,
          effectiveGpm: areaEntry.gpm * demandMultiplier,
          effectiveCfs: (areaEntry.gpm * demandMultiplier) / 448.831,
          effectiveArea: effectiveAreaForSizing,
          demandMultiplier: demandMultiplier,
          horizontal: horizontalResult,
          vertical: verticalResult,
          warningText: warningText
        });
      }

      setStormPreviewSummary(totalArea, gpmEffective, cfsEffective, areaResultRows.length, 'Live preview updates as you edit inputs. Click CALCULATE to commit to Results.', false);

      if (!commitResults) return;

      document.getElementById('selectedCodeText').innerHTML =
        '<div class="storm-kv-list">' +
        '<div class="storm-kv-row"><span class="storm-kv-label">Code Basis</span><span class="storm-kv-value">' + data.label + '</span></div>' +
        '<div class="storm-kv-row"><span class="storm-kv-label">Drain Type</span><span class="storm-kv-value">' + drainTypeLabel + '</span></div>' +
        '<div class="storm-kv-row"><span class="storm-kv-label">Rainfall Rate</span><span class="storm-kv-value">' + formatNumber(rainfallRate, 2) + ' in/hr</span></div>' +
        '</div>';
      document.getElementById('drainTypeDisplay').innerHTML = '';
      document.getElementById('rainfallDisplay').innerHTML = '';
      document.getElementById('totalAreaResult').innerHTML = formatNumber(totalArea, 0) + ' sq ft';
      document.getElementById('gpmText').innerHTML =
        '<div class="storm-metric-grid">' +
        '<div class="storm-metric"><span class="storm-metric-label">Gallons per Minute (GPM)</span><span class="storm-metric-value">' + formatNumber(gpmEffective, 2) + '</span></div>' +
        '<div class="storm-metric"><span class="storm-metric-label">Cubic Feet per Second (CFS)</span><span class="storm-metric-value">' + formatNumber(cfsEffective, 3) + '</span></div>' +
        '</div>';
      if (demandMultiplier > 1) {
        document.getElementById('cfsText').innerHTML = '<div class="storm-area-note">Base runoff: ' + formatNumber(gpmBase, 2) + ' GPM (' + formatNumber(cfsBase, 3) + ' CFS). Effective runoff uses ' + formatNumber(demandMultiplier, 2) + 'x demand.</div>';
      } else {
        document.getElementById('cfsText').innerHTML = '';
      }
      renderPerAreaResults(areaResultRows);
      updateStormResultPlaceholders();
      latestPerAreaExport = {
        codeBasisLabel: data.label,
        slopeLabel: slopeLabel,
        rainfallRate: formatNumber(rainfallRate, 2),
        drainTypeLabel: drainTypeLabel,
        demandMultiplier: formatNumber(demandMultiplier, 2) + 'x',
        calculatedAt: (new Date()).toLocaleString(),
        rows: []
      };
      for (i = 0; i < areaResultRows.length; i++) {
        latestPerAreaExport.rows.push({
          label: areaResultRows[i].label,
          parentArea: roundNumber(areaResultRows[i].parentArea, 0),
          subAreaCount: areaResultRows[i].subAreaCount,
          subAreaListText: areaResultRows[i].subAreaListText,
          roofArea: roundNumber(areaResultRows[i].roofArea, 0),
          sideWallCount: areaResultRows[i].sideWallCount,
          sideWallAreasText: (areaResultRows[i].sideWallCount > 0 ? areaResultRows[i].sideWallAreas.join(', ') : 'None'),
          sideWallSum: roundNumber(areaResultRows[i].sideWallSum, 0),
          sideWallFactorPercent: roundNumber(areaResultRows[i].sideWallFactor * 100, 0),
          sideWallContribution: roundNumber(areaResultRows[i].sideWallContribution, 0),
          area: roundNumber(areaResultRows[i].area, 0),
          effectiveArea: roundNumber(areaResultRows[i].effectiveArea, 0),
          baseGpm: roundNumber(areaResultRows[i].baseGpm, 2),
          effectiveGpm: roundNumber(areaResultRows[i].effectiveGpm, 2),
          effectiveCfs: roundNumber(areaResultRows[i].effectiveCfs, 3),
          horizontalSize: areaResultRows[i].horizontal.pipe.size,
          horizontalCapacity: roundNumber(areaResultRows[i].horizontal.pipe.gpm, 0),
          verticalSize: areaResultRows[i].vertical !== null ? areaResultRows[i].vertical.pipe.size : 'N/A',
          verticalCapacity: areaResultRows[i].vertical !== null ? roundNumber(areaResultRows[i].vertical.pipe.gpm, 0) : 'N/A',
          warningText: areaResultRows[i].warningText
        });
      }
      if (codeBasis == "IPC") {
        document.getElementById('basisNote').innerHTML = 'IPC per-area results are based on the vertical and horizontal GPM tables.';
      } else {
        document.getElementById('basisNote').innerHTML = 'CPC per-area results are based on the California vertical table and the California horizontal table you provided.';
      }
      if (demandMultiplier > 1) {
        document.getElementById('basisNote').innerHTML += ' Combined Storm mode applies a ' + formatNumber(demandMultiplier, 2) + 'x demand multiplier for sizing.';
      }
    }

    function resetResults() {
      document.getElementById('selectedCodeText').innerHTML = '';
      document.getElementById('drainTypeDisplay').innerHTML = '';
      document.getElementById('rainfallDisplay').innerHTML = '';
      document.getElementById('totalAreaResult').innerHTML = '';
      document.getElementById('gpmText').innerHTML = '';
      document.getElementById('cfsText').innerHTML = '';
      document.getElementById('perAreaResultsList').innerHTML = '';
      document.getElementById('basisNote').innerHTML = '';
      latestPerAreaExport = null;
      clearValidationSummary();
      updateStormResultPlaceholders();
    }

    function resetStormSection() {
      document.getElementById('areasContainer').innerHTML = makeAreaBlock(getAreaLabel(0), 0);
      document.getElementById('codeBasis').value = 'IPC';
      document.getElementById('drainType').value = 'PRIMARY';
      document.getElementById('zipCode').value = '';
      document.getElementById('rainfallRate').value = String(APP_DEFAULTS.stormRainfallRate || 1);
      setZipLookupStatus('', '');
      updateSlopeOptions();
      document.getElementById('slope').value = '1/8 in/ft';
      resetResults();
      setStormPreviewSummary(0, 0, 0, 1, 'Live preview updates as you edit inputs. Click CALCULATE to commit to Results.', false);
      recalculate(false);
    }

    function resetCondensateSection() {
      latestCondensateSnapshot = null;
      if (document.getElementById('condSizingMode')) document.getElementById('condSizingMode').value = 'PER_UNIT_QTY';
      condensateZones = [];
      condensateZoneIdSeq = 0;
      condensateRowIdSeq = 0;
      addCondensateZone({ id: makeCondensateZoneId(), name: 'Zone 1', equipmentRows: [{ id: makeCondensateRowId(), equipment: 'AHU', tonsPerUnit: 0, quantity: (APP_DEFAULTS.condensateQty || 1) }] });
      updateCondensatePreview();
      if (document.getElementById('condResults')) document.getElementById('condResults').innerHTML = '<span class="cond-results-placeholder">Condensate flow and pipe sizing results will appear here after calculation.</span>';
    }

    function resetVentSection() {
      latestVentSnapshot = null;
      if (document.getElementById('ventCodeBasis')) document.getElementById('ventCodeBasis').value = 'IPC';
      if (document.getElementById('ventUsageType')) document.getElementById('ventUsageType').value = 'PRIVATE';
      if (document.getElementById('ventDrainageOrientation')) document.getElementById('ventDrainageOrientation').value = 'HORIZONTAL';
      ventFixtureRows = [];
      ventFixtureIdSeq = 0;
      if (document.getElementById('ventUseManualTotal')) document.getElementById('ventUseManualTotal').checked = false;
      if (document.getElementById('ventManualTotalDfu')) document.getElementById('ventManualTotalDfu').value = 0;
      if (document.getElementById('ventAutoTotalDfu')) document.getElementById('ventAutoTotalDfu').value = 0;
      if (document.getElementById('ventTotalDfu')) document.getElementById('ventTotalDfu').value = 0;
      addVentFixtureRow();
      if (document.getElementById('ventPreview')) document.getElementById('ventPreview').innerHTML = APP_MESSAGES.calculateFirst || 'Preview updates as you edit inputs. Click CALCULATE to commit to Results.';
      if (document.getElementById('ventResults')) document.getElementById('ventResults').innerHTML = 'Enter fixture rows and press Calculate.';
    }

    function resetGasSection() {
      latestGasSnapshot = null;
      if (document.getElementById('gasCodeBasis')) document.getElementById('gasCodeBasis').value = 'IPC';
      if (document.getElementById('gasGasType')) document.getElementById('gasGasType').value = 'NATURAL_GAS';
      if (document.getElementById('gasMaterial')) document.getElementById('gasMaterial').value = 'SCHEDULE_40_METALLIC';
      refreshGasCriteriaDropdowns();
      if (document.getElementById('gasDemandMode')) document.getElementById('gasDemandMode').value = 'CFH';
      if (document.getElementById('gasHeatingValue')) document.getElementById('gasHeatingValue').value = 0;
      clearGasLineRows();
      addGasLineRow({ id: makeGasLineId(), label: 'Line 1', demandValue: 0, runLength: 0 });
      if (document.getElementById('gasPreview')) document.getElementById('gasPreview').innerHTML = APP_MESSAGES.calculateFirst || 'Preview updates as you edit inputs. Click CALCULATE to commit to Results.';
      if (document.getElementById('gasResults')) document.getElementById('gasResults').innerHTML = 'Enter gas inputs and press Calculate.';
    }

    function resetSolarSection() {
      latestSolarSnapshot = null;
      solarClimateManualOverride = false;
      if (solarZipLookupTimer) {
        window.clearTimeout(solarZipLookupTimer);
        solarZipLookupTimer = null;
      }
      if (solarAutoCalcTimer) {
        window.clearTimeout(solarAutoCalcTimer);
        solarAutoCalcTimer = null;
      }
      if (document.getElementById('solarBuildingSf')) document.getElementById('solarBuildingSf').value = 0;
      if (document.getElementById('solarZip')) document.getElementById('solarZip').value = '';
      if (document.getElementById('solarClimateZone')) document.getElementById('solarClimateZone').value = '';
      if (document.getElementById('solarUseSaraPath')) document.getElementById('solarUseSaraPath').checked = false;
      if (document.getElementById('solarSara')) document.getElementById('solarSara').value = '';
      if (document.getElementById('solarDValue')) document.getElementById('solarDValue').value = 0;
      if (document.getElementById('solarExempt')) document.getElementById('solarExempt').checked = false;
      if (document.getElementById('batteryExempt')) document.getElementById('batteryExempt').checked = false;
      if (document.getElementById('solarNotes')) document.getElementById('solarNotes').value = '';
      onSolarSaraToggle();
      updateSolarPreview();
      if (document.getElementById('solarResults')) document.getElementById('solarResults').innerHTML = 'Enter worksheet inputs to see live Solar results (or press CALCULATE).';
    }

    function resetDuctSection() {
      latestDuctSnapshot = null;
      if (ductLiveTimer) {
        window.clearTimeout(ductLiveTimer);
        ductLiveTimer = null;
      }
      if (document.getElementById('ductMode')) document.getElementById('ductMode').value = DUCT_CALC.modes.SIZE_FROM_CFM_FRICTION;
      if (document.getElementById('ductShape')) document.getElementById('ductShape').value = DUCT_CALC.shapes.ROUND;
      if (document.getElementById('ductCfm')) document.getElementById('ductCfm').value = 0;
      if (document.getElementById('ductFrictionTarget')) document.getElementById('ductFrictionTarget').value = 0;
      if (document.getElementById('ductVelocityTarget')) document.getElementById('ductVelocityTarget').value = 0;
      if (document.getElementById('ductDiameter')) document.getElementById('ductDiameter').value = 0;
      if (document.getElementById('ductWidth')) document.getElementById('ductWidth').value = 0;
      if (document.getElementById('ductHeight')) document.getElementById('ductHeight').value = 0;
      if (document.getElementById('ductRatioLimit')) document.getElementById('ductRatioLimit').value = 3;
      updateDuctModeUI();
      setDuctPreviewContent({}, APP_MESSAGES.ductPlaceholder || 'Enter airflow and sizing criteria to see live duct size recommendations.', '');
      setDuctResultPlaceholder(APP_MESSAGES.ductPlaceholder || 'Enter airflow and sizing criteria to see live duct size recommendations.');
    }

    function resetDuctStaticSection() {
      latestDuctStaticSnapshot = null;
      ductStaticSegments = [];
      ductStaticIdSeq = 1;
      addDuctStaticSegment({
        name: 'S1',
        lineType: 'Straight',
        lengthFt: 0,
        frictionRate: 0,
        width: 0,
        height: 0,
        radius: 0,
        velocityFpm: 0
      });
      if (document.getElementById('ductStaticResults')) {
        document.getElementById('ductStaticResults').innerHTML = '<div class="result-placeholder">' + (APP_MESSAGES.ductStaticPlaceholder || 'Enter duct segment data to calculate total static pressure drop.') + '</div>';
      }
    }

    function rawWsfuFixtureUnitByUsage(item, usageType) {
      if (!item) return null;
      var unit = (usageType === 'ASSEMBLY') ? item.assemblyWsfu : ((usageType === 'PUBLIC') ? item.publicWsfu : item.privateWsfu);
      return (unit === null || unit === undefined || isNaN(unit)) ? null : parseNonNegativeNumber(unit, 0);
    }

    function wsfuFixtureUnitByUsage(item, usageType) {
      var unit = rawWsfuFixtureUnitByUsage(item, usageType);
      return unit === null ? 0 : unit;
    }

    function getWsfuFixtureByKey(key) {
      var i;
      for (i = 0; i < wsfuFixtureCatalog.length; i++) {
        if (wsfuFixtureCatalog[i].key === key) return wsfuFixtureCatalog[i];
      }
      return wsfuFixtureCatalog.length > 0 ? wsfuFixtureCatalog[0] : null;
    }

    function normalizeCpcFixtureKey(key) {
      return String(key || '');
    }

    function computeWsfuRowsAndTotals() {
      var usageType = document.getElementById('wsfuUsageType') ? document.getElementById('wsfuUsageType').value : 'PRIVATE';
      var autoTotal = 0;
      var i, row, item, unit, quantity, factor, rowTotal;
      for (i = 0; i < wsfuFixtureRows.length; i++) {
        row = wsfuFixtureRows[i];
        item = getWsfuFixtureByKey(row.fixtureKey);
        unit = wsfuFixtureUnitByUsage(item, usageType);
        quantity = parseNonNegativeNumber(row.quantity, 0);
        factor = row.hotCold ? 0.75 : 1.0;
        row.unitWsfu = unit;
        row.rowTotalWsfu = roundNumber(quantity * unit * factor, 3);
        row.fixtureLabel = item ? item.label : '';
        autoTotal += row.rowTotalWsfu;
      }
      autoTotal = roundNumber(autoTotal, 3);
      if (document.getElementById('wsfuAutoTotalFu')) document.getElementById('wsfuAutoTotalFu').value = roundNumber(autoTotal, 3);
      return autoTotal;
    }

    function getWsfuFixtureRowsState() {
      var list = [];
      var i, row;
      for (i = 0; i < wsfuFixtureRows.length; i++) {
        row = wsfuFixtureRows[i];
        list.push({
          id: row.id,
          fixtureKey: normalizeCpcFixtureKey(row.fixtureKey),
          fixtureLabel: row.fixtureLabel || '',
          quantity: parseNonNegativeNumber(row.quantity, 0),
          hotCold: !!row.hotCold,
          unitWsfu: parseNonNegativeNumber(row.unitWsfu, 0),
          rowTotalWsfu: parseNonNegativeNumber(row.rowTotalWsfu, 0)
        });
      }
      return list;
    }

    function resetWsfuSection() {
      latestWsfuSnapshot = null;
      if (document.getElementById('wsfuFlushType')) document.getElementById('wsfuFlushType').value = 'FLUSH_TANK';
      if (document.getElementById('wsfuUsageType')) document.getElementById('wsfuUsageType').value = 'PRIVATE';
      if (document.getElementById('wsfuDesignLength')) document.getElementById('wsfuDesignLength').value = APP_DEFAULTS.wsfuDesignLengthFt || 100;
      if (document.getElementById('wsfuUseManualTotal')) document.getElementById('wsfuUseManualTotal').checked = false;
      if (document.getElementById('wsfuManualTotalFu')) document.getElementById('wsfuManualTotalFu').value = 0;
      wsfuFixtureRows = [];
      wsfuFixtureIdSeq = 1;
      addWsfuFixtureRow();
      updateWsfuTotalsAndUI();
      if (document.getElementById('wsfuResults')) document.getElementById('wsfuResults').innerHTML = 'Enter fixture rows and press Calculate.';
    }

    function resetForm() {
      resetStormSection();
      resetCondensateSection();
      resetVentSection();
      resetVentilationSection();
      resetGasSection();
      resetWsfuSection();
      resetFixtureUnitSection();
      resetDuctSection();
      resetDuctStaticSection();
      resetSolarSection();
      resetRefrigerantSection();
    }

    function condensateSizeFromTons(totalTonsUsed) {
      if (totalTonsUsed <= 20) return { size: '3/4"', warning: '' };
      if (totalTonsUsed <= 40) return { size: '1"', warning: '' };
      if (totalTonsUsed <= 90) return { size: '1-1/4"', warning: '' };
      if (totalTonsUsed <= 125) return { size: '1-1/2"', warning: '' };
      if (totalTonsUsed <= 250) return { size: '2"', warning: '' };
      return { size: 'OUT OF RANGE', warning: 'Total tons exceeds Table 814.3 maximum range (250 tons).' };
    }

    function updateCondensateModeUI() {
      updateCondensatePreview();
    }

    function makeCondensateZoneId() {
      condensateZoneIdSeq += 1;
      return 'zone_' + condensateZoneIdSeq;
    }

    function makeCondensateRowId() {
      condensateRowIdSeq += 1;
      return 'cond_row_' + condensateRowIdSeq;
    }

    function getCondensateEquipmentLabel(equipmentKey) {
      if (equipmentKey === 'FCU') return 'Fan Coil Unit (FCU)';
      if (equipmentKey === 'RTU') return 'Rooftop Unit (RTU)';
      if (equipmentKey === 'DEH') return 'Dehumidifier (DEH)';
      return 'Air Handler Unit (AHU)';
    }

    function addCondensateZone(seedZone) {
      var zone = {
        id: (seedZone && seedZone.id) ? seedZone.id : makeCondensateZoneId(),
        name: (seedZone && seedZone.name) ? String(seedZone.name) : ('Zone ' + (condensateZones.length + 1)),
        equipmentRows: []
      };
      condensateZones.push(zone);
      addCondensateEquipmentRow(zone.id, (seedZone && seedZone.equipmentRows && seedZone.equipmentRows.length > 0) ? seedZone.equipmentRows[0] : null, true);
      if (seedZone && seedZone.equipmentRows && seedZone.equipmentRows.length > 1) {
        var i;
        for (i = 1; i < seedZone.equipmentRows.length; i++) addCondensateEquipmentRow(zone.id, seedZone.equipmentRows[i], true);
      }
      renderCondensateZones();
      updateCondensatePreview();
    }

    function removeCondensateZone(zoneId) {
      var i;
      if (condensateZones.length <= 1) {
        alert('At least one zone is required.');
        return;
      }
      for (i = condensateZones.length - 1; i >= 0; i--) {
        if (condensateZones[i].id === zoneId) {
          condensateZones.splice(i, 1);
          break;
        }
      }
      renderCondensateZones();
      updateCondensatePreview();
    }

    function addCondensateEquipmentRow(zoneId, seedRow, deferRender) {
      var i, zone = null;
      for (i = 0; i < condensateZones.length; i++) {
        if (condensateZones[i].id === zoneId) {
          zone = condensateZones[i];
          break;
        }
      }
      if (!zone) return;
      zone.equipmentRows.push({
        id: (seedRow && seedRow.id) ? seedRow.id : makeCondensateRowId(),
        equipment: (seedRow && seedRow.equipment) ? seedRow.equipment : 'AHU',
        tonsPerUnit: parseNonNegativeNumber(seedRow && seedRow.tonsPerUnit, 0),
        quantity: parseNonNegativeNumber(seedRow && seedRow.quantity, 1)
      });
      if (!deferRender) {
        renderCondensateZones();
        updateCondensatePreview();
      }
    }

    function removeCondensateEquipmentRow(zoneId, rowId) {
      var i, j, zone = null;
      for (i = 0; i < condensateZones.length; i++) {
        if (condensateZones[i].id === zoneId) {
          zone = condensateZones[i];
          break;
        }
      }
      if (!zone) return;
      if (zone.equipmentRows.length <= 1) {
        alert('At least one equipment row is required per zone.');
        return;
      }
      for (j = zone.equipmentRows.length - 1; j >= 0; j--) {
        if (zone.equipmentRows[j].id === rowId) {
          zone.equipmentRows.splice(j, 1);
          break;
        }
      }
      renderCondensateZones();
      updateCondensatePreview();
    }

    function onCondensateZoneFieldChange(zoneId, field, rawValue) {
      var i;
      for (i = 0; i < condensateZones.length; i++) {
        if (condensateZones[i].id !== zoneId) continue;
        if (field === 'name') condensateZones[i].name = String(rawValue || '');
        break;
      }
      updateCondensatePreview();
    }

    function onCondensateRowFieldChange(zoneId, rowId, field, rawValue) {
      var i, j, zone, rows;
      for (i = 0; i < condensateZones.length; i++) {
        if (condensateZones[i].id !== zoneId) continue;
        zone = condensateZones[i];
        rows = zone.equipmentRows || [];
        for (j = 0; j < rows.length; j++) {
          if (rows[j].id !== rowId) continue;
          if (field === 'equipment') rows[j].equipment = rawValue || 'AHU';
          if (field === 'tonsPerUnit') rows[j].tonsPerUnit = parseNonNegativeNumber(rawValue, 0);
          if (field === 'quantity') rows[j].quantity = parseNonNegativeNumber(rawValue, 0);
          break;
        }
        break;
      }
      updateCondensatePreview();
    }

    function renderCondensateZones() {
      var container = document.getElementById('condZonesContainer');
      var html = '';
      var i, j, zone, row;
      if (!container) return;
      if (condensateZones.length <= 0) {
        addCondensateZone({ name: 'Zone 1', equipmentRows: [{ equipment: 'AHU', tonsPerUnit: 0, quantity: 1 }] });
        return;
      }
      for (i = 0; i < condensateZones.length; i++) {
        zone = condensateZones[i];
        html += '<div class="cond-zone-card">';
        html += '  <div class="cond-zone-header">';
        html += '    <div class="cond-zone-title">Zone ' + (i + 1) + '</div>';
        html += '    <div class="cond-zone-name-wrap">';
        html += '      <span class="cond-zone-name-label">Zone Name</span>';
        html += '      <input type="text" value="' + (zone.name || ('Zone ' + (i + 1))) + '" oninput="onCondensateZoneFieldChange(\'' + zone.id + '\', \'name\', this.value);" onchange="onCondensateZoneFieldChange(\'' + zone.id + '\', \'name\', this.value);" />';
        html += '    </div>';
        html += '  </div>';
        html += '  <div class="cond-zone-body">';
        for (j = 0; j < zone.equipmentRows.length; j++) {
          row = zone.equipmentRows[j];
          html += '    <div class="cond-equip-row">';
          html += '      <div class="cond-equip-main">';
          html += '        <div class="cond-equip-label">Equipment</div>';
          html += '        <div class="cond-equip-grid">';
          html += '          <div class="cond-equip-field"><span class="cond-equip-field-label">Equipment Type</span>';
          html += '            <select onchange="onCondensateRowFieldChange(\'' + zone.id + '\', \'' + row.id + '\', \'equipment\', this.value);">';
          html += '      <option value="AHU"' + (row.equipment === 'AHU' ? ' selected="selected"' : '') + '>Air Handler Unit (AHU)</option>';
          html += '      <option value="FCU"' + (row.equipment === 'FCU' ? ' selected="selected"' : '') + '>Fan Coil Unit (FCU)</option>';
          html += '      <option value="RTU"' + (row.equipment === 'RTU' ? ' selected="selected"' : '') + '>Rooftop Unit (RTU)</option>';
          html += '      <option value="DEH"' + (row.equipment === 'DEH' ? ' selected="selected"' : '') + '>Dehumidifier (DEH)</option>';
          html += '            </select>';
          html += '          </div>';
          html += '          <div class="cond-equip-field"><span class="cond-equip-field-label">Tons per Unit</span><input type="number" step="0.01" value="' + parseNonNegativeNumber(row.tonsPerUnit, 0) + '" oninput="onCondensateRowFieldChange(\'' + zone.id + '\', \'' + row.id + '\', \'tonsPerUnit\', this.value);" onchange="onCondensateRowFieldChange(\'' + zone.id + '\', \'' + row.id + '\', \'tonsPerUnit\', this.value);" /></div>';
          html += '          <div class="cond-equip-field"><span class="cond-equip-field-label">Number of Units</span><input type="number" step="1" value="' + parseNonNegativeNumber(row.quantity, 1) + '" oninput="onCondensateRowFieldChange(\'' + zone.id + '\', \'' + row.id + '\', \'quantity\', this.value);" onchange="onCondensateRowFieldChange(\'' + zone.id + '\', \'' + row.id + '\', \'quantity\', this.value);" /></div>';
          html += '        </div>';
          html += '      </div>';
          html += '      <div class="cond-equip-action"><input type="button" class="btn btn-small cond-danger-btn" value="Remove Equipment" onclick="removeCondensateEquipmentRow(\'' + zone.id + '\', \'' + row.id + '\');" /></div>';
          html += '    </div>';
        }
        html += '  </div>';
        html += '  <div class="cond-zone-actions">';
        html += '    <div class="cond-zone-actions-left">';
        html += '      <input type="button" class="btn cond-secondary-btn" value="+ Add Equipment" onclick="addCondensateEquipmentRow(\'' + zone.id + '\', { equipment: \'AHU\', tonsPerUnit: 0, quantity: 1 });" />';
        html += '      <input type="button" class="btn cond-danger-btn" value="Remove Zone" onclick="removeCondensateZone(\'' + zone.id + '\');" />';
        html += '    </div>';
        html += '  </div>';
        html += '</div>';
      }
      container.innerHTML = html;
    }

    function getCondensateZonesState() {
      var out = [];
      var i, j, zone, row, zoneRows;
      for (i = 0; i < condensateZones.length; i++) {
        zone = condensateZones[i];
        zoneRows = [];
        for (j = 0; j < zone.equipmentRows.length; j++) {
          row = zone.equipmentRows[j];
          zoneRows.push({
            id: row.id,
            equipment: row.equipment || 'AHU',
            tonsPerUnit: parseNonNegativeNumber(row.tonsPerUnit, 0),
            quantity: parseNonNegativeNumber(row.quantity, 0),
            rowTotalTons: roundNumber(parseNonNegativeNumber(row.tonsPerUnit, 0) * parseNonNegativeNumber(row.quantity, 0), 3)
          });
        }
        out.push({
          id: zone.id,
          name: zone.name || ('Zone ' + (i + 1)),
          equipmentRows: zoneRows
        });
      }
      return out;
    }

    function computeCondensateSnapshot() {
      var zones = getCondensateZonesState();
      var outZones = [];
      var i, j, zone, rows, row, rowTotal, rowSize, zoneTotal, zoneSize;
      if (zones.length <= 0) return { ok: false, reason: 'Add at least one zone.' };
      for (i = 0; i < zones.length; i++) {
        zone = zones[i];
        rows = zone.equipmentRows || [];
        if (rows.length <= 0) return { ok: false, reason: (zone.name || ('Zone ' + (i + 1))) + ': add at least one equipment row.' };
        zoneTotal = 0;
        var resultRows = [];
        for (j = 0; j < rows.length; j++) {
          row = rows[j];
          if (parseNonNegativeNumber(row.tonsPerUnit, 0) <= 0) return { ok: false, reason: (zone.name || ('Zone ' + (i + 1))) + ' row ' + (j + 1) + ': tons per unit must be greater than zero.' };
          if (parseNonNegativeNumber(row.quantity, 0) <= 0) return { ok: false, reason: (zone.name || ('Zone ' + (i + 1))) + ' row ' + (j + 1) + ': quantity must be greater than zero.' };
          rowTotal = parseNonNegativeNumber(row.tonsPerUnit, 0) * parseNonNegativeNumber(row.quantity, 0);
          rowSize = condensateSizeFromTons(rowTotal);
          resultRows.push({
            id: row.id,
            equipment: row.equipment || 'AHU',
            equipmentLabel: getCondensateEquipmentLabel(row.equipment || 'AHU'),
            tonsPerUnit: roundNumber(row.tonsPerUnit, 3),
            quantity: roundNumber(row.quantity, 3),
            rowTotalTons: roundNumber(rowTotal, 3),
            rowSize: rowSize.size,
            rowWarning: rowSize.warning || ''
          });
          zoneTotal += rowTotal;
        }
        zoneSize = condensateSizeFromTons(zoneTotal);
        outZones.push({
          id: zone.id,
          name: zone.name || ('Zone ' + (i + 1)),
          rows: resultRows,
          zoneTotalTons: roundNumber(zoneTotal, 3),
          recommendedSize: zoneSize.size,
          warning: zoneSize.warning || ''
        });
      }
      return {
        ok: true,
        tableName: 'TABLE 814.3',
        zones: outZones
      };
    }

    function updateCondensatePreview() {
      var node = document.getElementById('condPreview');
      var calc = computeCondensateSnapshot();
      if (!node) return;
      if (!calc.ok) {
        node.innerHTML = calc.reason;
        return;
      }
      var parts = [];
      var i;
      for (i = 0; i < calc.zones.length; i++) {
        parts.push((calc.zones[i].name || ('Zone ' + (i + 1))) + ': ' + formatNumber(calc.zones[i].zoneTotalTons, 2) + ' tons -> ' + calc.zones[i].recommendedSize);
      }
      node.innerHTML = 'Preview: ' + parts.join(' | ');
    }

    function renderCondensateResultHtml(snapshot) {
      var html = '';
      var zones = (snapshot.zones && snapshot.zones.length) ? snapshot.zones : [];
      var i, j, zone, rows;
      for (i = 0; i < zones.length; i++) {
        zone = zones[i];
        rows = zone.rows || [];
        html += '<div class="cond-result-zone">';
        html += '  <div class="cond-result-summary">';
        html += '    <div class="cond-result-metric cond-result-metric-zone"><span class="cond-result-metric-label">Zone</span><span class="cond-result-metric-value">' + (zone.name || ('Zone ' + (i + 1))) + '</span></div>';
        html += '    <div class="cond-result-metric"><span class="cond-result-metric-label">Total Tons</span><span class="cond-result-metric-value">' + formatNumber(parseNonNegativeNumber(zone.zoneTotalTons, 0), 2) + ' tons</span></div>';
        html += '    <div class="cond-result-metric cond-result-metric-main"><span class="cond-result-metric-label">Recommended Pipe Size</span><span class="cond-result-metric-value">' + (zone.recommendedSize || '') + '</span></div>';
        html += '  </div>';
        html += '  <div class="cond-result-table-wrap">';
        html += '    <table class="cond-result-table">';
        html += '      <colgroup>';
        html += '        <col class="col-eq" />';
        html += '        <col class="col-tons" />';
        html += '        <col class="col-qty" />';
        html += '        <col class="col-rowtons" />';
        html += '        <col class="col-rowsize" />';
        html += '      </colgroup>';
        html += '      <thead><tr><th>Equipment</th><th class="num">Tons/Unit</th><th class="num">Qty</th><th class="num">Row Tons</th><th class="num">Row Size</th></tr></thead>';
        html += '      <tbody>';
        for (j = 0; j < rows.length; j++) {
          html += '        <tr>';
          html += '          <td>' + (rows[j].equipmentLabel || getCondensateEquipmentLabel(rows[j].equipment)) + '</td>';
          html += '          <td class="num">' + formatNumber(parseNonNegativeNumber(rows[j].tonsPerUnit, 0), 2) + '</td>';
          html += '          <td class="num">' + formatNumber(parseNonNegativeNumber(rows[j].quantity, 0), 0) + '</td>';
          html += '          <td class="num">' + formatNumber(parseNonNegativeNumber(rows[j].rowTotalTons, 0), 2) + '</td>';
          html += '          <td class="num">' + (rows[j].rowSize || '') + '</td>';
          html += '        </tr>';
        }
        html += '      </tbody>';
        html += '    </table>';
        html += '  </div>';
        html += '  <div class="cond-result-final">';
        html += '    <span class="cond-result-final-label">Recommended Minimum Pipe Size</span>';
        html += '    <span class="cond-result-final-value">' + (zone.recommendedSize || '') + '</span>';
        if (zone.warning) html += '<div class="cond-result-warning">' + zone.warning + '</div>';
        html += '  </div>';
        html += '</div>';
      }
      html += '<div class="cond-result-footer-note">Sizing based on ' + (snapshot.tableName || 'TABLE 814.3') + ' ranges (tons of refrigeration).</div>';
      return html;
    }

    function calculateCondensate() {
      var calc = computeCondensateSnapshot();
      var resultNode = document.getElementById('condResults');
      if (!calc.ok) {
        latestCondensateSnapshot = null;
        setWarningResult(resultNode, calc.reason);
        return;
      }
      latestCondensateSnapshot = {
        section: 'condensate',
        tableName: calc.tableName,
        sizingMode: 'PER_UNIT_QTY',
        zones: calc.zones,
        rows: [],
        totalTonsUsed: 0,
        recommendedSize: '',
        warning: ''
      };
      if (resultNode) resultNode.innerHTML = renderCondensateResultHtml(latestCondensateSnapshot);
    }

    function makeVentFixtureId() {
      ventFixtureIdSeq += 1;
      return 'vent_fixture_' + ventFixtureIdSeq;
    }

    function getVentFixtureByKey(key) {
      var i;
      for (i = 0; i < ventFixtureCatalog.length; i++) {
        if (ventFixtureCatalog[i].key === key) return ventFixtureCatalog[i];
      }
      return ventFixtureCatalog.length > 0 ? ventFixtureCatalog[0] : null;
    }

    function getVentUsageType() {
      var node = document.getElementById('ventUsageType');
      return (node && node.value === 'ASSEMBLY') ? 'ASSEMBLY' : ((node && node.value === 'PUBLIC') ? 'PUBLIC' : 'PRIVATE');
    }

    function getVentDrainageOrientation() {
      var node = document.getElementById('ventDrainageOrientation');
      return (node && node.value === 'VERTICAL') ? 'VERTICAL' : 'HORIZONTAL';
    }

    function findSanitary7032ByDfu(dfu, orientation) {
      var rows = sanitary7032Table[(orientation === 'VERTICAL') ? 'VERTICAL' : 'HORIZONTAL'] || [];
      var i, rowMax;
      for (i = 0; i < rows.length; i++) {
        rowMax = parseNonNegativeNumber(rows[i].maxUnits, 0);
        if (dfu <= rowMax) return rows[i];
      }
      return null;
    }

    function ventFixtureDfuByUsage(item, usageType) {
      if (!item) return null;
      var raw = (usageType === 'ASSEMBLY') ? item.assemblyDfu : ((usageType === 'PUBLIC') ? item.publicDfu : item.privateDfu);
      if (raw === null || raw === undefined || raw === '' || isNaN(parseFloat(raw))) return null;
      return parseNonNegativeNumber(raw, 0);
    }

    function addVentFixtureRow(seed) {
      var defaultFixture = ventFixtureCatalog.length > 0 ? ventFixtureCatalog[0].key : '';
      var usageType = getVentUsageType();
      var i, candidate;
      for (i = 0; i < ventFixtureCatalog.length; i++) {
        candidate = ventFixtureCatalog[i];
        if (ventFixtureDfuByUsage(candidate, usageType) !== null) {
          defaultFixture = candidate.key;
          break;
        }
      }
      ventFixtureRows.push({
        id: (seed && seed.id) ? seed.id : makeVentFixtureId(),
        fixtureKey: (seed && seed.fixtureKey) ? seed.fixtureKey : defaultFixture,
        quantity: parseNonNegativeNumber((seed && seed.quantity !== undefined) ? seed.quantity : 0, 0),
        unitDfu: parseNonNegativeNumber((seed && seed.unitDfu !== undefined) ? seed.unitDfu : 0, 0),
        rowTotalDfu: parseNonNegativeNumber((seed && seed.rowTotalDfu !== undefined) ? seed.rowTotalDfu : 0, 0),
        fixtureLabel: (seed && seed.fixtureLabel) ? seed.fixtureLabel : ''
      });
      renderVentFixtureRows();
      updateVentTotalsAndUI();
    }

    function removeVentFixtureRow(id) {
      var i;
      for (i = 0; i < ventFixtureRows.length; i++) {
        if (ventFixtureRows[i].id === id) {
          ventFixtureRows.splice(i, 1);
          break;
        }
      }
      if (ventFixtureRows.length <= 0) addVentFixtureRow();
      else {
        renderVentFixtureRows();
        updateVentTotalsAndUI();
      }
    }

    function onVentFixtureFieldChange(id, field, value) {
      var i, row;
      for (i = 0; i < ventFixtureRows.length; i++) {
        if (ventFixtureRows[i].id !== id) continue;
        row = ventFixtureRows[i];
        if (field === 'fixtureKey') row.fixtureKey = normalizeCpcFixtureKey(value);
        if (field === 'quantity') row.quantity = parseNonNegativeNumber(value, 0);
        break;
      }
      updateVentTotalsAndUI();
    }

    function renderVentFixtureRows() {
      var node = document.getElementById('ventFixtureRows');
      var html = '';
      var i, j, row, item, options, usageType, unit;
      if (!node) return;
      usageType = getVentUsageType();
      html += '<table class="vent-fixture-table">';
      html += '<colgroup><col class="col-fixture" /><col class="col-qty" /><col class="col-unit" /><col class="col-row" /><col class="col-action" /></colgroup>';
      html += '<thead><tr><th>Fixture</th><th class="num">Qty</th><th class="num">Unit DFU</th><th class="num">Row DFU</th><th></th></tr></thead><tbody>';
      for (i = 0; i < ventFixtureRows.length; i++) {
        row = ventFixtureRows[i];
        options = '';
        for (j = 0; j < ventFixtureCatalog.length; j++) {
          item = ventFixtureCatalog[j];
          unit = ventFixtureDfuByUsage(item, usageType);
          if (unit === null) continue;
          options += '<option value="' + item.key + '"' + (item.key === normalizeCpcFixtureKey(row.fixtureKey) ? ' selected' : '') + '>' + item.label + '</option>';
        }
        if (!options) options = '<option value="">No fixtures available</option>';
        html += '<tr>';
        html += '<td><select style="width:100%;" onchange="onVentFixtureFieldChange(\'' + row.id + '\', \'fixtureKey\', this.value);">' + options + '</select></td>';
        html += '<td class="num"><input type="number" step="1" min="0" class="vent-qty-input" value="' + parseNonNegativeNumber(row.quantity, 0) + '" oninput="onVentFixtureFieldChange(\'' + row.id + '\', \'quantity\', this.value);" /></td>';
        html += '<td class="num"><span class="vent-dfu-badge">' + formatNumber(parseNonNegativeNumber(row.unitDfu, 0), 2) + '</span></td>';
        html += '<td class="num"><span class="vent-dfu-badge">' + formatNumber(parseNonNegativeNumber(row.rowTotalDfu, 0), 2) + '</span></td>';
        html += '<td class="num"><input type="button" class="btn btn-small vent-remove-btn" value="Remove" onclick="removeVentFixtureRow(\'' + row.id + '\');" /></td>';
        html += '</tr>';
      }
      html += '</tbody></table>';
      node.innerHTML = html;
    }

    function computeVentRowsAndTotals() {
      var i, row, item, unit, autoTotal = 0;
      var usageType = getVentUsageType();
      var firstAvailableKey = '';
      for (i = 0; i < ventFixtureCatalog.length; i++) {
        if (ventFixtureDfuByUsage(ventFixtureCatalog[i], usageType) !== null) {
          firstAvailableKey = ventFixtureCatalog[i].key;
          break;
        }
      }
      for (i = 0; i < ventFixtureRows.length; i++) {
        row = ventFixtureRows[i];
        item = getVentFixtureByKey(row.fixtureKey);
        unit = ventFixtureDfuByUsage(item, usageType);
        if (unit === null && firstAvailableKey) {
          row.fixtureKey = firstAvailableKey;
          item = getVentFixtureByKey(row.fixtureKey);
          unit = ventFixtureDfuByUsage(item, usageType);
        }
        if (unit === null) unit = 0;
        row.unitDfu = unit;
        row.rowTotalDfu = roundNumber(parseNonNegativeNumber(row.quantity, 0) * parseNonNegativeNumber(unit, 0), 3);
        row.fixtureLabel = item ? item.label : '';
        autoTotal += row.rowTotalDfu;
      }
      autoTotal = roundNumber(autoTotal, 3);
      if (document.getElementById('ventAutoTotalDfu')) document.getElementById('ventAutoTotalDfu').value = autoTotal;
      return autoTotal;
    }

    function getVentFixtureRowsState() {
      var out = [];
      var i, row;
      for (i = 0; i < ventFixtureRows.length; i++) {
        row = ventFixtureRows[i];
        out.push({
          id: row.id,
          fixtureKey: normalizeCpcFixtureKey(row.fixtureKey),
          fixtureLabel: row.fixtureLabel || '',
          quantity: parseNonNegativeNumber(row.quantity, 0),
          unitDfu: parseNonNegativeNumber(row.unitDfu, 0),
          rowTotalDfu: parseNonNegativeNumber(row.rowTotalDfu, 0)
        });
      }
      return out;
    }

    function updateVentTotalsAndUI() {
      var autoTotal = computeVentRowsAndTotals();
      var useManual = !!(document.getElementById('ventUseManualTotal') && document.getElementById('ventUseManualTotal').checked);
      var usageType = getVentUsageType();
      var manualNode = document.getElementById('ventManualTotalDfu');
      var effectiveNode = document.getElementById('ventTotalDfu');
      var effectiveTotal = useManual ? parseNonNegativeNumber(manualNode ? manualNode.value : 0, 0) : autoTotal;
      if (manualNode) manualNode.disabled = !useManual;
      if (effectiveNode) effectiveNode.value = roundNumber(effectiveTotal, 3);
      renderVentFixtureRows();
      updateVentPreview();
    }

    function updateVentPreview() {
      var node = document.getElementById('ventPreview');
      if (!node) return;
      var codeBasis = document.getElementById('ventCodeBasis').value;
      var usageType = getVentUsageType();
      var orientation = getVentDrainageOrientation();
      var orientationLabel = (orientation === 'VERTICAL') ? 'Drainage Vertical' : 'Drainage Horizontal';
      var dfu = parseNonNegativeNumber(document.getElementById('ventTotalDfu') ? document.getElementById('ventTotalDfu').value : 0, 0);
      var match;
      if (isNaN(dfu) || dfu <= 0) {
        node.innerHTML = '<div class="vent-preview-strip"><span class="vent-preview-item"><span class="vent-preview-label">Status</span><span class="vent-preview-value">Add sanitary fixtures and enter DFU > 0</span></span></div>';
        return;
      }
      match = findSanitary7032ByDfu(dfu, orientation);
      if (!match) {
        node.innerHTML =
          '<div class="vent-preview-strip">' +
          '<span class="vent-preview-item"><span class="vent-preview-label">Code Basis</span><span class="vent-preview-value">' + codeBasis + '</span></span>' +
          '<span class="vent-preview-item"><span class="vent-preview-label">Orientation</span><span class="vent-preview-value">' + orientationLabel + '</span></span>' +
          '<span class="vent-preview-item"><span class="vent-preview-label">Usage</span><span class="vent-preview-value">' + usageType + '</span></span>' +
          '<span class="vent-preview-item"><span class="vent-preview-label">Effective DFU</span><span class="vent-preview-value">' + formatNumber(dfu, 2) + '</span></span>' +
          '<span class="vent-preview-item"><span class="vent-preview-label">Selected Pipe Size</span><span class="vent-preview-value">Out of range</span></span>' +
          '</div>';
        return;
      }
      node.innerHTML =
        '<div class="vent-preview-strip">' +
        '<span class="vent-preview-item"><span class="vent-preview-label">Code Basis</span><span class="vent-preview-value">' + codeBasis + '</span></span>' +
        '<span class="vent-preview-item"><span class="vent-preview-label">Orientation</span><span class="vent-preview-value">' + orientationLabel + '</span></span>' +
        '<span class="vent-preview-item"><span class="vent-preview-label">Usage</span><span class="vent-preview-value">' + usageType + '</span></span>' +
        '<span class="vent-preview-item"><span class="vent-preview-label">Effective DFU</span><span class="vent-preview-value">' + formatNumber(dfu, 2) + '</span></span>' +
        '<span class="vent-preview-item"><span class="vent-preview-label">Selected Pipe Size</span><span class="vent-preview-value">' + match.size + '</span></span>' +
        '</div>';
    }

    function renderVentResultHtml(snapshot) {
      var rows = (snapshot && snapshot.rows && snapshot.rows.length) ? snapshot.rows : [];
      var i, row, rowsHtml = '';
      var orientation = (snapshot && snapshot.drainageOrientation === 'VERTICAL') ? 'VERTICAL' : 'HORIZONTAL';
      var orientationLabel = (orientation === 'VERTICAL') ? 'Drainage Vertical' : 'Drainage Horizontal';
      var warningHtml = (snapshot && snapshot.warning) ? ('<div class="warning-line">' + snapshot.warning + '</div>') : '';
      for (i = 0; i < rows.length; i++) {
        row = rows[i];
        rowsHtml += '<tr>' +
          '<td>' + (row.fixtureLabel || row.fixtureKey || '') + '</td>' +
          '<td class="num">' + formatNumber(parseNonNegativeNumber(row.quantity, 0), 0) + '</td>' +
          '<td class="num">' + formatNumber(parseNonNegativeNumber(row.unitDfu, 0), 2) + '</td>' +
          '<td class="num">' + formatNumber(parseNonNegativeNumber(row.rowTotalDfu, 0), 2) + '</td>' +
          '</tr>';
      }
      if (!rowsHtml) rowsHtml = '<tr><td colspan="4" class="subtle" style="padding:7px 8px;">No fixture rows.</td></tr>';
      return '' +
        '<div class="vent-result-summary">' +
        '<div class="vent-result-metric"><span class="vent-result-metric-label">Effective DFU</span><span class="vent-result-metric-value">' + formatNumber(parseNonNegativeNumber(snapshot.dfu, 0), 2) + '</span></div>' +
        '<div class="vent-result-metric vent-result-metric-main"><span class="vent-result-metric-label">Recommended Pipe Size</span><span class="vent-result-metric-value">' + (snapshot.recommendedSize || 'Out of range') + '</span></div>' +
        '<div class="vent-result-metric"><span class="vent-result-metric-label">Max Units</span><span class="vent-result-metric-value">' + (snapshot.tableMaxUnits ? formatNumber(parseNonNegativeNumber(snapshot.tableMaxUnits, 0), 0) : 'N/A') + '</span></div>' +
        '<div class="vent-result-metric"><span class="vent-result-metric-label">Max Length</span><span class="vent-result-metric-value">' + (snapshot.maxLengthReference || 'N/A') + '</span></div>' +
        '</div>' +
        warningHtml +
        '<div class="vent-result-table-wrap">' +
        '<table class="vent-result-table">' +
        '<thead><tr><th>Fixture</th><th class="num">Qty</th><th class="num">Unit DFU</th><th class="num">Row DFU</th></tr></thead><tbody>' +
        rowsHtml +
        '</tbody></table></div>' +
        '<div class="vent-result-note">Code Basis: ' + (snapshot.codeBasis || 'IPC') + ' | Usage: ' + fixtureUnitCategoryLabel(snapshot.usageType) + ' | Orientation: ' + orientationLabel + ' | Sizing Table: ' + (snapshot.sizingTable || 'TABLE 703.2') + '.</div>';
    }

    function calculateVent() {
      var codeBasis = document.getElementById('ventCodeBasis').value;
      var usageType = getVentUsageType();
      var drainageOrientation = getVentDrainageOrientation();
      var dfu = parseNonNegativeNumber(document.getElementById('ventTotalDfu') ? document.getElementById('ventTotalDfu').value : 0, 0);
      var useManual = !!(document.getElementById('ventUseManualTotal') && document.getElementById('ventUseManualTotal').checked);
      var autoTotal = parseNonNegativeNumber(document.getElementById('ventAutoTotalDfu') ? document.getElementById('ventAutoTotalDfu').value : 0, 0);
      var manualTotal = parseNonNegativeNumber(document.getElementById('ventManualTotalDfu') ? document.getElementById('ventManualTotalDfu').value : 0, 0);
      var rowsState = getVentFixtureRowsState();
      var resultNode = document.getElementById('ventResults');
      var sizeMatch = null;
      var warningText = '';
      var maxLengthText = '';

      if (isNaN(dfu) || dfu <= 0) {
        latestVentSnapshot = null;
        setWarningResult(resultNode, 'Enter an effective DFU value greater than zero.');
        return;
      }

      sizeMatch = findSanitary7032ByDfu(dfu, drainageOrientation);
      if (!sizeMatch) {
        warningText = 'Effective DFU exceeds Table 703.2 maximum for ' + ((drainageOrientation === 'VERTICAL') ? 'Drainage Vertical' : 'Drainage Horizontal') + '.';
      } else {
        maxLengthText = (sizeMatch.maxLength === '' || sizeMatch.maxLength === null || sizeMatch.maxLength === undefined) ? 'N/A' : String(sizeMatch.maxLength);
      }
      latestVentSnapshot = {
        section: 'vent',
        codeBasis: codeBasis,
        usageType: usageType,
        drainageOrientation: drainageOrientation,
        dfu: roundNumber(dfu, 3),
        recommendedSize: sizeMatch ? sizeMatch.size : '',
        sizingTable: 'TABLE 703.2',
        tableMaxUnits: sizeMatch ? parseNonNegativeNumber(sizeMatch.maxUnits, 0) : 0,
        maxLengthReference: sizeMatch ? maxLengthText : 'N/A',
        warning: warningText,
        autoTotalDfu: roundNumber(autoTotal, 3),
        manualTotalDfu: roundNumber(manualTotal, 3),
        useManualTotal: useManual,
        rows: rowsState
      };
      if (resultNode) resultNode.innerHTML = renderVentResultHtml(latestVentSnapshot);
    }

    function makeWsfuFixtureId() {
      wsfuFixtureIdSeq += 1;
      return 'fixture_' + wsfuFixtureIdSeq;
    }

    function addWsfuFixtureRow(seed) {
      var defaultFixture = wsfuFixtureCatalog.length > 0 ? wsfuFixtureCatalog[0].key : '';
      var row = {
        id: (seed && seed.id) ? seed.id : makeWsfuFixtureId(),
        fixtureKey: (seed && seed.fixtureKey) ? seed.fixtureKey : defaultFixture,
        quantity: parseNonNegativeNumber((seed && seed.quantity !== undefined) ? seed.quantity : 0, 0),
        hotCold: !!(seed && seed.hotCold),
        unitWsfu: parseNonNegativeNumber((seed && seed.unitWsfu !== undefined) ? seed.unitWsfu : 0, 0),
        rowTotalWsfu: parseNonNegativeNumber((seed && seed.rowTotalWsfu !== undefined) ? seed.rowTotalWsfu : 0, 0),
        fixtureLabel: (seed && seed.fixtureLabel) ? seed.fixtureLabel : ''
      };
      wsfuFixtureRows.push(row);
      renderWsfuFixtureRows();
      updateWsfuTotalsAndUI();
    }

    function removeWsfuFixtureRow(id) {
      var i;
      for (i = 0; i < wsfuFixtureRows.length; i++) {
        if (wsfuFixtureRows[i].id === id) {
          wsfuFixtureRows.splice(i, 1);
          break;
        }
      }
      if (wsfuFixtureRows.length <= 0) addWsfuFixtureRow();
      else {
        renderWsfuFixtureRows();
        updateWsfuTotalsAndUI();
      }
    }

    function onWsfuFixtureFieldChange(rowId, field, value) {
      var i, row;
      for (i = 0; i < wsfuFixtureRows.length; i++) {
        row = wsfuFixtureRows[i];
        if (row.id !== rowId) continue;
        if (field === 'fixtureKey') row.fixtureKey = value;
        if (field === 'quantity') row.quantity = parseNonNegativeNumber(value, 0);
        if (field === 'hotCold') row.hotCold = !!value;
        break;
      }
      updateWsfuTotalsAndUI();
    }

    function renderWsfuFixtureRows() {
      var node = document.getElementById('wsfuFixtureRows');
      if (!node) return;
      var usageType = document.getElementById('wsfuUsageType') ? document.getElementById('wsfuUsageType').value : 'PRIVATE';
      var i, j, row, item, unit, options, disabledTag, html;
      html = '<table class="wsfu-fixture-table">' +
        '<thead><tr>' +
        '<th>Fixture</th>' +
        '<th class="num">Qty</th>' +
        '<th class="center">Hot+Cold</th>' +
        '<th class="num">Unit WSFU</th>' +
        '<th class="num">Row Total</th>' +
        '<th class="num">Action</th>' +
        '</tr></thead><tbody>';
      for (i = 0; i < wsfuFixtureRows.length; i++) {
        row = wsfuFixtureRows[i];
        options = '';
        for (j = 0; j < wsfuFixtureCatalog.length; j++) {
          item = wsfuFixtureCatalog[j];
          unit = rawWsfuFixtureUnitByUsage(item, usageType);
          disabledTag = unit === null ? ' (N/A)' : '';
          options += '<option value="' + item.key + '"' + (item.key === normalizeCpcFixtureKey(row.fixtureKey) ? ' selected' : '') + '>' + item.label + disabledTag + '</option>';
        }
        html += '<tr>' +
          '<td><select class="wsfu-fixture-select" onchange="onWsfuFixtureFieldChange(\'' + row.id + '\', \'fixtureKey\', this.value);">' + options + '</select></td>' +
          '<td class="num"><input type="number" min="0" step="1" value="' + parseNonNegativeNumber(row.quantity, 0) + '" class="wsfu-qty-input" onchange="onWsfuFixtureFieldChange(\'' + row.id + '\', \'quantity\', this.value);" /></td>' +
          '<td class="center"><input type="checkbox" class="wsfu-hotcold-checkbox" ' + (row.hotCold ? 'checked' : '') + ' onchange="onWsfuFixtureFieldChange(\'' + row.id + '\', \'hotCold\', this.checked);" /></td>' +
          '<td class="num wsfu-cell-metric">' + formatNumber(parseNonNegativeNumber(row.unitWsfu, 0), 3) + '</td>' +
          '<td class="num wsfu-cell-metric">' + formatNumber(parseNonNegativeNumber(row.rowTotalWsfu, 0), 3) + '</td>' +
          '<td class="num"><input type="button" class="btn btn-remove-row wsfu-remove-btn" value="Remove" onclick="removeWsfuFixtureRow(\'' + row.id + '\');" /></td>' +
          '</tr>';
      }
      html += '</tbody></table>';
      node.innerHTML = html;
    }

    function updateWsfuTotalsAndUI() {
      var autoTotal = computeWsfuRowsAndTotals();
      var useManual = !!(document.getElementById('wsfuUseManualTotal') && document.getElementById('wsfuUseManualTotal').checked);
      var manualNode = document.getElementById('wsfuManualTotalFu');
      var effectiveNode = document.getElementById('wsfuTotalFu');
      var manualValue = parseNonNegativeNumber(manualNode ? manualNode.value : 0, 0);
      if (manualNode) manualNode.disabled = !useManual;
      if (effectiveNode) effectiveNode.value = roundNumber((useManual ? manualValue : autoTotal), 3);
      renderWsfuFixtureRows();
      if (document.getElementById('wsfuPreview')) {
        var basis = document.getElementById('wsfuFlushType') ? document.getElementById('wsfuFlushType').value : 'FLUSH_TANK';
        var length = parseNonNegativeNumber(document.getElementById('wsfuDesignLength') ? document.getElementById('wsfuDesignLength').value : (APP_DEFAULTS.wsfuDesignLengthFt || 100), APP_DEFAULTS.wsfuDesignLengthFt || 100);
        document.getElementById('wsfuPreview').innerHTML =
          'Preview: Effective Total WSFU ' + formatNumber(parseNonNegativeNumber(effectiveNode ? effectiveNode.value : 0, 0), 3) +
          ' | Basis ' + (basis === 'FLUSH_VALVE' ? 'Flush Valve' : 'Flush Tank') +
          ' | Design Length ' + formatNumber(length, 2) + ' ft';
      }
    }

    function renderWsfuResult(response) {
      var rows = (response && response.frictionLossByPipeType && response.frictionLossByPipeType.length) ? response.frictionLossByPipeType : [];
      var sizingBasis = (response && response.flushType === 'FLUSH_VALVE') ? 'Flush Valve' : 'Flush Tank';
      var selectedTotalGpm = (response && response.flushType === 'FLUSH_VALVE') ? parseNonNegativeNumber(response.flushValveGpm, 0) : parseNonNegativeNumber(response.flushTankGpm, 0);
      var selectedMatchedFu = (response && response.flushType === 'FLUSH_VALVE') ? parseNonNegativeNumber(response.matchedValveFu, 0) : parseNonNegativeNumber(response.matchedTankFu, 0);
      var minPsi = Infinity;
      var minIdx = -1;
      var i;
      for (i = 0; i < rows.length; i++) {
        var psi = parseNonNegativeNumber(rows[i].psiLossForLength, 0);
        if (psi < minPsi) {
          minPsi = psi;
          minIdx = i;
        }
      }

      var html = '' +
        '<div class="wsfu-results-section">' +
          '<div class="wsfu-results-subtitle">Results Summary</div>' +
          '<div class="wsfu-results-grid">' +
            '<div class="wsfu-result-item"><span class="wsfu-result-label">Total Fixture Units</span><span class="wsfu-result-value">' + formatNumber(parseNonNegativeNumber(response.totalFu, 0), 2) + '</span></div>' +
            '<div class="wsfu-result-item"><span class="wsfu-result-label">Sizing Basis</span><span class="wsfu-result-value">' + sizingBasis + '</span></div>' +
            '<div class="wsfu-result-item"><span class="wsfu-result-label">Design Length</span><span class="wsfu-result-value">' + formatNumber(parseNonNegativeNumber(response.designLengthFt, 0), 2) + ' ft</span></div>' +
            '<div class="wsfu-result-item"><span class="wsfu-result-label">Selected Flow</span><span class="wsfu-result-value">' + formatNumber(parseNonNegativeNumber(response.selectedFlowGpm, 0), 2) + ' GPM</span></div>' +
            '<div class="wsfu-result-item"><span class="wsfu-result-label">Total GPM (' + sizingBasis + ')</span><span class="wsfu-result-value">' + formatNumber(selectedTotalGpm, 2) + '</span></div>' +
            '<div class="wsfu-result-item"><span class="wsfu-result-label">Matched FU (' + sizingBasis + ')</span><span class="wsfu-result-value">' + formatNumber(selectedMatchedFu, 2) + '</span></div>' +
          '</div>' +
          '<div class="wsfu-selected-size"><span class="wsfu-selected-size-label">Selected Water Service Size</span><span class="wsfu-selected-size-value">' + (response.serviceSize || '') + '</span><span class="wsfu-selected-size-note">Capacity: ' + formatNumber(parseNonNegativeNumber(response.sizeCapacityGpm, 0), 2) + ' GPM</span></div>' +
        '</div>';

      if (response.warning) {
        html += '<div class="warning-line">' + response.warning + '</div>';
      }
      if (rows.length > 0) {
        html += '<div class="wsfu-results-section">' +
          '<div class="wsfu-results-subtitle">Friction Loss by Pipe Type</div>' +
          '<table class="wsfu-friction-table">' +
          '<thead><tr>' +
          '<th>Pipe Type</th>' +
          '<th class="num">C</th>' +
          '<th class="num">ft/100ft</th>' +
          '<th class="num">psi/100ft</th>' +
          '<th class="num">psi @ length</th>' +
          '</tr></thead><tbody>';
        for (i = 0; i < rows.length; i++) {
          html += '<tr class="' + (i === minIdx ? 'wsfu-friction-best' : '') + '">' +
            '<td>' + (rows[i].pipeType || '') + '</td>' +
            '<td class="num">' + formatNumber(parseNonNegativeNumber(rows[i].cFactor, 0), 0) + '</td>' +
            '<td class="num">' + formatNumber(parseNonNegativeNumber(rows[i].headLossFtPer100, 0), 3) + '</td>' +
            '<td class="num">' + formatNumber(parseNonNegativeNumber(rows[i].psiLossPer100, 0), 3) + '</td>' +
            '<td class="num">' + formatNumber(parseNonNegativeNumber(rows[i].psiLossForLength, 0), 3) + '</td>' +
            '</tr>';
        }
        html += '</tbody></table></div>';
      }
      html += '<div class="module-note wsfu-result-source">Source: ' + (response.source || 'WSFU2GPM workbook') + '</div>';
      return html;
    }

    async function calculateWsfu() {
      updateWsfuTotalsAndUI();
      var totalFuNode = document.getElementById('wsfuTotalFu');
      var flushTypeNode = document.getElementById('wsfuFlushType');
      var designLengthNode = document.getElementById('wsfuDesignLength');
      var resultNode = document.getElementById('wsfuResults');
      var totalFu = parseFloat(totalFuNode ? totalFuNode.value : '');
      var designLength = parseFloat(designLengthNode ? designLengthNode.value : '');
      var flushType = (flushTypeNode && flushTypeNode.value === 'FLUSH_VALVE') ? 'FLUSH_VALVE' : 'FLUSH_TANK';
      var wsfuSelection = buildProjectPayload().wsfu.selection;
      if (isNaN(totalFu) || totalFu < 0) {
        latestWsfuSnapshot = null;
        setWarningResult(resultNode, 'Enter a non-negative total fixture unit value.');
        return;
      }
      if (isNaN(designLength) || designLength < 0) {
        latestWsfuSnapshot = null;
        setWarningResult(resultNode, 'Enter a non-negative design length.');
        return;
      }
      if (!window.pywebview || !window.pywebview.api || !window.pywebview.api.calculate_wsfu) {
        latestWsfuSnapshot = null;
        setWarningResult(resultNode, 'WSFU service is not ready yet. Please try again.');
        return;
      }
      try {
        var response = await window.pywebview.api.calculate_wsfu({
          totalFu: totalFu,
          flushType: flushType,
          designLengthFt: designLength
        });
        if (!response || !response.ok) {
          latestWsfuSnapshot = null;
          setWarningResult(resultNode, (response && response.reason ? response.reason : 'WSFU lookup failed.'));
          return;
        }
        latestWsfuSnapshot = {
          section: 'wsfu',
          selection: wsfuSelection,
          result: response
        };
        if (resultNode) {
          resultNode.innerHTML = renderWsfuResult(response);
        }
      } catch (e) {
        latestWsfuSnapshot = null;
        setWarningResult(resultNode, 'WSFU lookup failed. ' + ((e && e.message) ? e.message : 'Unknown error'));
      }
    }

    function updateDuctModeUI() {
      var mode = document.getElementById('ductMode') ? normalizeDuctMode(document.getElementById('ductMode').value, document.getElementById('ductShape') ? document.getElementById('ductShape').value : DUCT_CALC.shapes.ROUND) : DUCT_CALC.modes.SIZE_FROM_CFM_FRICTION;
      if (document.getElementById('ductMode')) document.getElementById('ductMode').value = mode;
      var shape = document.getElementById('ductShape') ? document.getElementById('ductShape').value : DUCT_CALC.shapes.ROUND;
      var frictionRow = document.getElementById('ductFrictionRow');
      var velocityRow = document.getElementById('ductVelocityRow');
      var roundRow = document.getElementById('ductRoundRow');
      var rectRow = document.getElementById('ductRectRow');
      var shapeRow = document.getElementById('ductShapeRow');
      var frictionActive = (mode === DUCT_CALC.modes.SIZE_FROM_CFM_FRICTION);
      var velocityActive = (mode === DUCT_CALC.modes.SIZE_FROM_CFM_VELOCITY);
      var roundActive = (mode === DUCT_CALC.modes.CHECK_ROUND_SIZE);
      var rectActive = (mode === DUCT_CALC.modes.CHECK_RECT_SIZE);
      var shapeActive = (mode === DUCT_CALC.modes.SIZE_FROM_CFM_FRICTION || mode === DUCT_CALC.modes.SIZE_FROM_CFM_VELOCITY);
      if (mode === DUCT_CALC.modes.CHECK_ROUND_SIZE && document.getElementById('ductShape')) document.getElementById('ductShape').value = DUCT_CALC.shapes.ROUND;
      if (mode === DUCT_CALC.modes.CHECK_RECT_SIZE && document.getElementById('ductShape')) document.getElementById('ductShape').value = DUCT_CALC.shapes.RECTANGULAR;
      var rows = [
        { node: frictionRow, active: frictionActive },
        { node: velocityRow, active: velocityActive },
        { node: roundRow, active: roundActive },
        { node: rectRow, active: rectActive }
      ];
      var i, row, inputs, j;
      for (i = 0; i < rows.length; i++) {
        row = rows[i];
        if (!row.node) continue;
        row.node.classList.toggle('duct-row-inactive', !row.active);
        inputs = row.node.querySelectorAll('input, select');
        for (j = 0; j < inputs.length; j++) inputs[j].disabled = !row.active;
      }
      if (shapeRow) {
        shapeRow.classList.toggle('duct-row-inactive', !shapeActive);
        var shapeButtons = shapeRow.querySelectorAll('button');
        for (j = 0; j < shapeButtons.length; j++) shapeButtons[j].disabled = !shapeActive;
      }
      syncDuctSegmentedControls();
    }

    function onDuctInputsChanged() {
      updateDuctModeUI();
      if (ductLiveTimer) window.clearTimeout(ductLiveTimer);
      ductLiveTimer = window.setTimeout(function () {
        calculateDuct(false);
      }, 120);
    }

    function setDuctPreviewContent(result, statusText, statusClass) {
      var node = document.getElementById('ductPreview');
      if (!node) return;
      node.innerHTML =
        '<div class="duct-preview-item"><span class="duct-preview-label">Eq. Diameter</span><span class="duct-preview-value">' + (result.equivalentDiameter ? (formatNumber(parseNonNegativeNumber(result.equivalentDiameter, 0), 3) + ' in') : '-') + '</span></div>' +
        '<div class="duct-preview-item"><span class="duct-preview-label">Round</span><span class="duct-preview-value">' + (result.recommendedRound || '-') + '</span></div>' +
        '<div class="duct-preview-item"><span class="duct-preview-label">Rect</span><span class="duct-preview-value">' + (result.recommendedRect || '-') + '</span></div>' +
        '<div class="duct-preview-item"><span class="duct-preview-label">Velocity</span><span class="duct-preview-value">' + (result.velocityFpm ? (formatNumber(parseNonNegativeNumber(result.velocityFpm, 0), 1) + ' FPM') : '-') + '</span></div>' +
        '<div class="duct-preview-item"><span class="duct-preview-label">Friction</span><span class="duct-preview-value">' + (result.frictionRate ? (formatNumber(parseNonNegativeNumber(result.frictionRate, 0), 4) + ' in.wg/100ft') : '-') + '</span></div>' +
        '<div class="duct-preview-status ' + (statusClass || '') + '">' + (statusText || '') + '</div>';
    }

    function renderDuctResult(result) {
      var html = '';
      var velocityTarget = parseNonNegativeNumber((result && result.velocityTarget) ? result.velocityTarget : 0, 0);
      html += '<div class="duct-result-summary">' +
        '<div class="duct-result-metric"><span class="duct-result-metric-label">Mode</span><span class="duct-result-metric-value">' + (result.modeLabel || getDuctModeLabel(DUCT_CALC.modes.SIZE_FROM_CFM_FRICTION)) + '</span></div>' +
        '<div class="duct-result-metric"><span class="duct-result-metric-label">Equivalent Diameter</span><span class="duct-result-metric-value">' + formatNumber(parseNonNegativeNumber(result.equivalentDiameter, 0), 3) + ' in</span></div>' +
        '<div class="duct-result-metric duct-result-metric-main"><span class="duct-result-metric-label">Recommended Round Size</span><span class="duct-result-metric-value">' + (result.recommendedRound || 'N/A') + '</span></div>' +
        '<div class="duct-result-metric duct-result-metric-main"><span class="duct-result-metric-label">Recommended Rectangular Size</span><span class="duct-result-metric-value">' + (result.recommendedRect || 'N/A') + '</span></div>' +
        '</div>';
      html += '<div class="duct-result-table-wrap"><table class="duct-result-table"><thead><tr>' +
        '<th>Performance Metric</th><th class="num">Value</th>' +
        '</tr></thead><tbody>' +
        '<tr><td>Velocity</td><td class="num">' + formatNumber(parseNonNegativeNumber(result.velocityFpm, 0), 1) + ' FPM</td></tr>' +
        (velocityTarget > 0 ? ('<tr><td>Entered Velocity Target</td><td class="num">' + formatNumber(velocityTarget, 1) + ' FPM</td></tr>') : '') +
        '<tr><td>Velocity Pressure</td><td class="num">' + formatNumber(parseNonNegativeNumber(result.velocityPressure, 0), 4) + ' in.wg</td></tr>' +
        '<tr><td>Friction Rate</td><td class="num">' + formatNumber(parseNonNegativeNumber(result.frictionRate, 0), 4) + ' in.wg/100ft</td></tr>' +
        '</tbody></table></div>';
      if (result.warning) html += '<div class="warning-line">' + result.warning + '</div>';
      html += '<div class="duct-result-note">' + (result.message || '') + '</div>';
      html += '<div class="duct-result-note">Calibration note: Baseline formulas are active; tune with DuctSizer sample cases for exact parity.</div>';
      return html;
    }

    function setDuctResultPlaceholder(message) {
      var node = document.getElementById('ductResults');
      setPlaceholderResult(node, message);
    }

    function calculateDuct(commitResults) {
      if (commitResults !== true) commitResults = false;
      var sel = getDuctSelectionState();
      var node = document.getElementById('ductResults');
      var mode = sel.mode;
      var shape = sel.shape;
      var cfm = sel.cfm;
      var eqDiameter = 0;
      var roundSize = '';
      var rectSize = '';
      var velocity = 0;
      var vp = 0;
      var friction = 0;
      var warning = '';
      var targetRound, chosenRound, chosenRect;
      var velocityTarget = parseNonNegativeNumber(sel.velocityTarget, 0);
      var ratioLimit = parseNonNegativeNumber(sel.ratioLimit, 3);
      var checkShape = (mode === DUCT_CALC.modes.CHECK_RECT_SIZE) ? DUCT_CALC.shapes.RECTANGULAR : DUCT_CALC.shapes.ROUND;

      if (cfm <= 0) {
        setDuctPreviewContent({}, 'Enter airflow (CFM) greater than zero.', 'warn');
        latestDuctSnapshot = null;
        setDuctResultPlaceholder(APP_MESSAGES.ductPlaceholder || 'Enter airflow and sizing criteria to see live duct size recommendations.');
        return;
      }

      if (mode === DUCT_CALC.modes.SIZE_FROM_CFM_FRICTION) {
        if (sel.frictionTarget <= 0) {
          setDuctPreviewContent({}, 'Enter friction target greater than zero for sizing mode.', 'warn');
          latestDuctSnapshot = null;
          setDuctResultPlaceholder(APP_MESSAGES.ductPlaceholder || 'Enter airflow and sizing criteria to see live duct size recommendations.');
          return;
        }
        targetRound = ductRoundSizeFromTarget(cfm, sel.frictionTarget);
        if (velocityTarget > 0) {
          var velocityDiameter = Math.sqrt((576 * cfm) / (Math.PI * velocityTarget));
          if (!isNaN(velocityDiameter) && velocityDiameter > targetRound) targetRound = velocityDiameter;
        }
        chosenRound = ductChooseStandardRound(targetRound);
        roundSize = chosenRound + ' in';
        eqDiameter = chosenRound;
        chosenRect = ductChooseRectFromEquivalent(eqDiameter, ratioLimit);
        if (chosenRect) rectSize = chosenRect.width + ' x ' + chosenRect.height + ' in';
        velocity = ductVelocityFpm(cfm, DUCT_CALC.shapes.ROUND, chosenRound, 0, 0);
        vp = ductVelocityPressure(velocity);
        friction = ductFrictionRate(cfm, eqDiameter);
      } else if (mode === DUCT_CALC.modes.SIZE_FROM_CFM_VELOCITY) {
        if (velocityTarget <= 0) {
          setDuctPreviewContent({}, 'Enter velocity target greater than zero for velocity sizing mode.', 'warn');
          latestDuctSnapshot = null;
          setDuctResultPlaceholder(APP_MESSAGES.ductPlaceholder || 'Enter airflow and sizing criteria to see live duct size recommendations.');
          return;
        }
        targetRound = Math.sqrt((576 * cfm) / (Math.PI * velocityTarget));
        chosenRound = ductChooseStandardRound(targetRound);
        roundSize = chosenRound + ' in';
        eqDiameter = chosenRound;
        chosenRect = ductChooseRectFromEquivalent(eqDiameter, ratioLimit);
        if (chosenRect) rectSize = chosenRect.width + ' x ' + chosenRect.height + ' in';
        velocity = ductVelocityFpm(cfm, DUCT_CALC.shapes.ROUND, chosenRound, 0, 0);
        vp = ductVelocityPressure(velocity);
        friction = ductFrictionRate(cfm, eqDiameter);
      } else if (mode === DUCT_CALC.modes.CHECK_ROUND_SIZE) {
        if (sel.diameter <= 0) {
          setDuctPreviewContent({}, 'Enter round diameter greater than zero.', 'warn');
          latestDuctSnapshot = null;
          setDuctResultPlaceholder(APP_MESSAGES.ductPlaceholder || 'Enter airflow and sizing criteria to see live duct size recommendations.');
          return;
        }
        eqDiameter = sel.diameter;
        roundSize = sel.diameter + ' in';
        chosenRect = ductChooseRectFromEquivalent(eqDiameter, ratioLimit);
        if (chosenRect) rectSize = chosenRect.width + ' x ' + chosenRect.height + ' in';
        velocity = ductVelocityFpm(cfm, DUCT_CALC.shapes.ROUND, sel.diameter, 0, 0);
        vp = ductVelocityPressure(velocity);
        friction = ductFrictionRate(cfm, eqDiameter);
      } else {
        if (sel.width <= 0 || sel.height <= 0) {
          setDuctPreviewContent({}, 'Enter rectangular width and height greater than zero.', 'warn');
          latestDuctSnapshot = null;
          setDuctResultPlaceholder(APP_MESSAGES.ductPlaceholder || 'Enter airflow and sizing criteria to see live duct size recommendations.');
          return;
        }
        eqDiameter = ductEquivalentDiameterIn(DUCT_CALC.shapes.RECTANGULAR, 0, sel.width, sel.height);
        chosenRound = ductChooseStandardRound(eqDiameter);
        roundSize = chosenRound + ' in (eq)';
        rectSize = sel.width + ' x ' + sel.height + ' in';
        velocity = ductVelocityFpm(cfm, DUCT_CALC.shapes.RECTANGULAR, 0, sel.width, sel.height);
        vp = ductVelocityPressure(velocity);
        friction = ductFrictionRate(cfm, eqDiameter);
      }

      if (checkShape === DUCT_CALC.shapes.RECTANGULAR && sel.height > 0 && sel.width / sel.height > sel.ratioLimit) {
        warning = 'Entered width/height exceeds recommended rectangular ratio.';
      }
      if (velocityTarget > 0 && velocity > velocityTarget) {
        warning = (warning ? (warning + ' ') : '') + 'Calculated velocity exceeds the entered velocity target.';
      }
      if (velocity > 3000) {
        warning = (warning ? (warning + ' ') : '') + 'Velocity is high; verify noise and pressure-drop criteria.';
      }

      var result = {
        status: 'calculated',
        modeLabel: getDuctModeLabel(mode),
        message: (mode === DUCT_CALC.modes.SIZE_FROM_CFM_FRICTION)
          ? 'Sized from airflow + friction target using standard round sizes and equivalent rectangular selection.'
          : (mode === DUCT_CALC.modes.SIZE_FROM_CFM_VELOCITY)
            ? 'Sized from airflow + velocity target using standard round sizes and equivalent rectangular selection.'
            : 'Checked using entered duct geometry and airflow.',
        equivalentDiameter: roundNumber(eqDiameter, 3),
        recommendedRound: roundSize,
        recommendedRect: rectSize,
        velocityFpm: roundNumber(velocity, 1),
        velocityPressure: roundNumber(vp, 4),
        frictionRate: roundNumber(friction, 4),
        velocityTarget: roundNumber(velocityTarget, 1),
        warning: warning
      };
      if (mode === DUCT_CALC.modes.SIZE_FROM_CFM_FRICTION || mode === DUCT_CALC.modes.SIZE_FROM_CFM_VELOCITY) {
        if (document.getElementById('ductDiameter')) document.getElementById('ductDiameter').value = roundNumber(eqDiameter, 3);
        if (chosenRect) {
          if (document.getElementById('ductWidth')) document.getElementById('ductWidth').value = chosenRect.width;
          if (document.getElementById('ductHeight')) document.getElementById('ductHeight').value = chosenRect.height;
        }
      } else if (mode === DUCT_CALC.modes.CHECK_ROUND_SIZE) {
        if (chosenRect) {
          if (document.getElementById('ductWidth')) document.getElementById('ductWidth').value = chosenRect.width;
          if (document.getElementById('ductHeight')) document.getElementById('ductHeight').value = chosenRect.height;
        }
      } else if (mode === DUCT_CALC.modes.CHECK_RECT_SIZE) {
        if (document.getElementById('ductDiameter')) document.getElementById('ductDiameter').value = roundNumber(eqDiameter, 3);
      }
      setDuctPreviewContent(result, warning ? warning : 'Live sizing summary is up to date.', warning ? 'warn' : 'ok');
      latestDuctSnapshot = { section: 'duct', selection: sel, result: result };
      if (node) node.innerHTML = renderDuctResult(result);
    }

    function makeDuctStaticId() {
      ductStaticIdSeq += 1;
      return 'ds_' + ductStaticIdSeq;
    }

    function nearestBin(value, bins) {
      var i, best = bins[0], bestDiff = Math.abs(value - bins[0]), d;
      for (i = 1; i < bins.length; i++) {
        d = Math.abs(value - bins[i]);
        if (d < bestDiff) {
          bestDiff = d;
          best = bins[i];
        }
      }
      return best;
    }

    function binIndex(value, bins) {
      var i;
      for (i = 0; i < bins.length; i++) {
        if (bins[i] === value) return i;
      }
      return 0;
    }

    function lookupElbowC(hwRatio, rwRatio) {
      var hwNearest = nearestBin(hwRatio, DUCT_STATIC_C_TABLE.hwBins);
      var rwNearest = nearestBin(rwRatio, DUCT_STATIC_C_TABLE.rwBins);
      var r = binIndex(rwNearest, DUCT_STATIC_C_TABLE.rwBins);
      var h = binIndex(hwNearest, DUCT_STATIC_C_TABLE.hwBins);
      return parseNonNegativeNumber(DUCT_STATIC_C_TABLE.values[r][h], 0);
    }

    function addDuctStaticSegment(seed) {
      ductStaticSegments.push({
        id: (seed && seed.id) ? seed.id : makeDuctStaticId(),
        name: (seed && seed.name) ? seed.name : ('S' + (ductStaticSegments.length + 1)),
        lineType: (seed && seed.lineType) ? seed.lineType : 'Straight',
        lengthFt: parseNonNegativeNumber((seed && seed.lengthFt !== undefined) ? seed.lengthFt : 0, 0),
        frictionRate: parseNonNegativeNumber((seed && seed.frictionRate !== undefined) ? seed.frictionRate : 0, 0),
        width: parseNonNegativeNumber((seed && seed.width !== undefined) ? seed.width : 0, 0),
        height: parseNonNegativeNumber((seed && seed.height !== undefined) ? seed.height : 0, 0),
        radius: parseNonNegativeNumber((seed && seed.radius !== undefined) ? seed.radius : 0, 0),
        velocityFpm: parseNonNegativeNumber((seed && seed.velocityFpm !== undefined) ? seed.velocityFpm : 0, 0),
        rwRatio: parseNonNegativeNumber((seed && seed.rwRatio !== undefined) ? seed.rwRatio : 0, 0),
        hwRatio: parseNonNegativeNumber((seed && seed.hwRatio !== undefined) ? seed.hwRatio : 0, 0),
        cCoefficient: parseNonNegativeNumber((seed && seed.cCoefficient !== undefined) ? seed.cCoefficient : 0, 0),
        segmentPressureDrop: parseNonNegativeNumber((seed && seed.segmentPressureDrop !== undefined) ? seed.segmentPressureDrop : 0, 0),
        warning: (seed && seed.warning) ? seed.warning : ''
      });
      renderDuctStaticSegments();
    }

    function removeDuctStaticSegment(id) {
      var i;
      for (i = 0; i < ductStaticSegments.length; i++) {
        if (ductStaticSegments[i].id === id) {
          ductStaticSegments.splice(i, 1);
          break;
        }
      }
      if (ductStaticSegments.length <= 0) addDuctStaticSegment();
      renderDuctStaticSegments();
    }

    function onDuctStaticSegmentChange(id, field, value) {
      var i, seg;
      for (i = 0; i < ductStaticSegments.length; i++) {
        seg = ductStaticSegments[i];
        if (seg.id !== id) continue;
        if (field === 'name') seg.name = String(value || '');
        else if (field === 'lineType') seg.lineType = String(value || 'Straight');
        else if (field === 'lengthFt') seg.lengthFt = parseNonNegativeNumber(value, 0);
        else if (field === 'frictionRate') seg.frictionRate = parseNonNegativeNumber(value, 0);
        else if (field === 'width') seg.width = parseNonNegativeNumber(value, 0);
        else if (field === 'height') seg.height = parseNonNegativeNumber(value, 0);
        else if (field === 'radius') seg.radius = parseNonNegativeNumber(value, 0);
        else if (field === 'velocityFpm') seg.velocityFpm = parseNonNegativeNumber(value, 0);
        break;
      }
      renderDuctStaticSegments();
    }

    function isDuctStaticFieldActive(lineType, fieldName) {
      if (lineType === 'Straight') return fieldName === 'lengthFt' || fieldName === 'frictionRate';
      if (lineType === 'Elbow') return fieldName === 'width' || fieldName === 'height' || fieldName === 'radius' || fieldName === 'velocityFpm';
      if (lineType === 'Transition' || lineType === 'Outlet') return false;
      return true;
    }

    function computeDuctStaticSegment(seg) {
      var width = parseNonNegativeNumber(seg.width, 0);
      var height = parseNonNegativeNumber(seg.height, 0);
      var radius = parseNonNegativeNumber(seg.radius, 0);
      var velocity = parseNonNegativeNumber(seg.velocityFpm, 0);
      var rw = (width > 0) ? (radius / width) : 0;
      var hw = (width > 0) ? (height / width) : 0;
      var c = 0;
      var drop = 0;
      var warning = '';
      if (seg.lineType === 'Straight') {
        drop = (parseNonNegativeNumber(seg.lengthFt, 0) * parseNonNegativeNumber(seg.frictionRate, 0)) / 100.0;
      } else if (seg.lineType === 'Elbow') {
        c = lookupElbowC(hw, rw);
        drop = Math.pow(velocity / 4000.0, 2) * c;
      } else {
        warning = seg.lineType + ' line is included in v1 as placeholder (0 drop until formula table is provided).';
        drop = 0;
      }
      seg.rwRatio = roundNumber(rw, 3);
      seg.hwRatio = roundNumber(hw, 3);
      seg.cCoefficient = roundNumber(c, 3);
      seg.segmentPressureDrop = roundNumber(drop, 5);
      seg.warning = warning;
      return seg;
    }

    function getDuctStaticSelectionState() {
      var out = [];
      var i, seg;
      for (i = 0; i < ductStaticSegments.length; i++) {
        seg = computeDuctStaticSegment(ductStaticSegments[i]);
        out.push({
          id: seg.id,
          name: seg.name,
          lineType: seg.lineType,
          lengthFt: parseNonNegativeNumber(seg.lengthFt, 0),
          frictionRate: parseNonNegativeNumber(seg.frictionRate, 0),
          width: parseNonNegativeNumber(seg.width, 0),
          height: parseNonNegativeNumber(seg.height, 0),
          radius: parseNonNegativeNumber(seg.radius, 0),
          velocityFpm: parseNonNegativeNumber(seg.velocityFpm, 0),
          rwRatio: parseNonNegativeNumber(seg.rwRatio, 0),
          hwRatio: parseNonNegativeNumber(seg.hwRatio, 0),
          cCoefficient: parseNonNegativeNumber(seg.cCoefficient, 0),
          segmentPressureDrop: parseNonNegativeNumber(seg.segmentPressureDrop, 0),
          warning: seg.warning || ''
        });
      }
      return { segments: out };
    }

    function renderDuctStaticSegments() {
      var node = document.getElementById('ductStaticSegments');
      if (!node) return;
      var html = '<div class="ductstatic-table-wrap"><table class="ductstatic-table">' +
        '<thead><tr>' +
        '<th>Segment</th>' +
        '<th>Type</th>' +
        '<th class="num">Length (ft)</th>' +
        '<th class="num">FR (in.wg/100ft)</th>' +
        '<th class="num">W (in)</th>' +
        '<th class="num">H (in)</th>' +
        '<th class="num">R (in)</th>' +
        '<th class="num">Velocity (FPM)</th>' +
        '<th class="num">R/W</th>' +
        '<th class="num">H/W</th>' +
        '<th class="num">C</th>' +
        '<th class="num">Drop (in.wg)</th>' +
        '<th class="num">Action</th>' +
        '</tr></thead><tbody>';
      var i, seg;
      for (i = 0; i < ductStaticSegments.length; i++) {
        seg = computeDuctStaticSegment(ductStaticSegments[i]);
        var activeLength = isDuctStaticFieldActive(seg.lineType, 'lengthFt');
        var activeFriction = isDuctStaticFieldActive(seg.lineType, 'frictionRate');
        var activeWidth = isDuctStaticFieldActive(seg.lineType, 'width');
        var activeHeight = isDuctStaticFieldActive(seg.lineType, 'height');
        var activeRadius = isDuctStaticFieldActive(seg.lineType, 'radius');
        var activeVelocity = isDuctStaticFieldActive(seg.lineType, 'velocityFpm');
        html += '<tr>' +
          '<td><input type="text" value="' + (seg.name || '') + '" class="input-seg" onchange="onDuctStaticSegmentChange(\'' + seg.id + '\', \'name\', this.value);" /></td>' +
          '<td><select class="input-type ductstatic-type-select" onchange="onDuctStaticSegmentChange(\'' + seg.id + '\', \'lineType\', this.value);">' +
            '<option value="Straight"' + (seg.lineType === 'Straight' ? ' selected' : '') + '>Straight</option>' +
            '<option value="Elbow"' + (seg.lineType === 'Elbow' ? ' selected' : '') + '>Elbow</option>' +
            '<option value="Transition"' + (seg.lineType === 'Transition' ? ' selected' : '') + '>Transition</option>' +
            '<option value="Outlet"' + (seg.lineType === 'Outlet' ? ' selected' : '') + '>Outlet</option>' +
          '</select></td>' +
          '<td class="num ' + (activeLength ? '' : 'ductstatic-cell-muted') + '"><input type="number" step="0.01" value="' + parseNonNegativeNumber(seg.lengthFt, 0) + '" class="input-num" ' + (activeLength ? '' : 'disabled="disabled"') + ' onchange="onDuctStaticSegmentChange(\'' + seg.id + '\', \'lengthFt\', this.value);" /></td>' +
          '<td class="num ' + (activeFriction ? '' : 'ductstatic-cell-muted') + '"><input type="number" step="0.0001" value="' + parseNonNegativeNumber(seg.frictionRate, 0) + '" class="input-num" ' + (activeFriction ? '' : 'disabled="disabled"') + ' onchange="onDuctStaticSegmentChange(\'' + seg.id + '\', \'frictionRate\', this.value);" /></td>' +
          '<td class="num ' + (activeWidth ? '' : 'ductstatic-cell-muted') + '"><input type="number" step="0.01" value="' + parseNonNegativeNumber(seg.width, 0) + '" class="input-num" ' + (activeWidth ? '' : 'disabled="disabled"') + ' onchange="onDuctStaticSegmentChange(\'' + seg.id + '\', \'width\', this.value);" /></td>' +
          '<td class="num ' + (activeHeight ? '' : 'ductstatic-cell-muted') + '"><input type="number" step="0.01" value="' + parseNonNegativeNumber(seg.height, 0) + '" class="input-num" ' + (activeHeight ? '' : 'disabled="disabled"') + ' onchange="onDuctStaticSegmentChange(\'' + seg.id + '\', \'height\', this.value);" /></td>' +
          '<td class="num ' + (activeRadius ? '' : 'ductstatic-cell-muted') + '"><input type="number" step="0.01" value="' + parseNonNegativeNumber(seg.radius, 0) + '" class="input-num" ' + (activeRadius ? '' : 'disabled="disabled"') + ' onchange="onDuctStaticSegmentChange(\'' + seg.id + '\', \'radius\', this.value);" /></td>' +
          '<td class="num ' + (activeVelocity ? '' : 'ductstatic-cell-muted') + '"><input type="number" step="0.01" value="' + parseNonNegativeNumber(seg.velocityFpm, 0) + '" class="input-num" ' + (activeVelocity ? '' : 'disabled="disabled"') + ' onchange="onDuctStaticSegmentChange(\'' + seg.id + '\', \'velocityFpm\', this.value);" /></td>' +
          '<td class="num">' + formatNumber(parseNonNegativeNumber(seg.rwRatio, 0), 3) + '</td>' +
          '<td class="num">' + formatNumber(parseNonNegativeNumber(seg.hwRatio, 0), 3) + '</td>' +
          '<td class="num">' + formatNumber(parseNonNegativeNumber(seg.cCoefficient, 0), 3) + '</td>' +
          '<td class="num">' + formatNumber(parseNonNegativeNumber(seg.segmentPressureDrop, 0), 5) + '</td>' +
          '<td class="remove-cell"><input type="button" class="btn btn-small ductstatic-remove-btn" value="Remove" onclick="removeDuctStaticSegment(\'' + seg.id + '\');" /></td>' +
          '</tr>';
        if (seg.warning) {
          html += '<tr><td colspan="13"><div class="ductstatic-warning-inline">' + seg.warning + '</div></td></tr>';
        }
      }
      html += '</tbody></table></div>';
      node.innerHTML = html;
      var previewNode = document.getElementById('ductStaticPreview');
      if (previewNode) {
        var sel = getDuctStaticSelectionState();
        var total = 0;
        var hasMeaningfulInputs = false;
        for (i = 0; i < sel.segments.length; i++) total += parseNonNegativeNumber(sel.segments[i].segmentPressureDrop, 0);
        for (i = 0; i < sel.segments.length; i++) {
          if (sel.segments[i].lineType === 'Straight' && parseNonNegativeNumber(sel.segments[i].lengthFt, 0) > 0 && parseNonNegativeNumber(sel.segments[i].frictionRate, 0) > 0) hasMeaningfulInputs = true;
          if (sel.segments[i].lineType === 'Elbow' &&
              parseNonNegativeNumber(sel.segments[i].velocityFpm, 0) > 0 &&
              parseNonNegativeNumber(sel.segments[i].width, 0) > 0 &&
              parseNonNegativeNumber(sel.segments[i].height, 0) > 0 &&
              parseNonNegativeNumber(sel.segments[i].radius, 0) > 0) hasMeaningfulInputs = true;
        }
        previewNode.innerHTML =
          '<div class="ductstatic-preview-item"><span class="ductstatic-preview-label">Segments</span><span class="ductstatic-preview-value">' + sel.segments.length + '</span></div>' +
          '<div class="ductstatic-preview-item"><span class="ductstatic-preview-label">Total Pressure Drop</span><span class="ductstatic-preview-value">' + formatNumber(total, 5) + ' in.wg</span></div>' +
          '<div class="ductstatic-preview-status ' + (hasMeaningfulInputs ? 'ok' : 'warn') + '">' + (hasMeaningfulInputs ? 'Segment worksheet is ready for final calculation.' : (APP_MESSAGES.ductStaticPlaceholder || 'Enter duct segment data to calculate total static pressure drop.')) + '</div>';
      }
    }

    function renderDuctStaticResult(result) {
      var html = '';
      var segmentCount = (result && result.segments) ? result.segments.length : 0;
      var i;
      html += '<div class="ductstatic-result-summary">' +
        '<div class="ductstatic-result-metric"><span class="ductstatic-result-metric-label">Segments</span><span class="ductstatic-result-metric-value">' + segmentCount + '</span></div>' +
        '<div class="ductstatic-result-metric ductstatic-result-metric-main"><span class="ductstatic-result-metric-label">Total Pressure Drop</span><span class="ductstatic-result-metric-value">' + formatNumber(parseNonNegativeNumber(result.totalPressureDrop, 0), 5) + ' in.wg</span></div>' +
        '</div>';
      html += '<div class="ductstatic-table-wrap"><table class="ductstatic-table"><thead><tr>' +
        '<th>Segment</th><th>Type</th><th class="num">Drop (in.wg)</th><th class="num">Velocity (FPM)</th><th class="num">FR (in.wg/100ft)</th>' +
        '</tr></thead><tbody>';
      for (i = 0; i < segmentCount; i++) {
        html += '<tr>' +
          '<td>' + (result.segments[i].name || ('S' + (i + 1))) + '</td>' +
          '<td>' + (result.segments[i].lineType || '') + '</td>' +
          '<td class="num">' + formatNumber(parseNonNegativeNumber(result.segments[i].segmentPressureDrop, 0), 5) + '</td>' +
          '<td class="num">' + formatNumber(parseNonNegativeNumber(result.segments[i].velocityFpm, 0), 1) + '</td>' +
          '<td class="num">' + formatNumber(parseNonNegativeNumber(result.segments[i].frictionRate, 0), 4) + '</td>' +
          '</tr>';
      }
      html += '</tbody></table></div>';
      if (result.warnings && result.warnings.length) {
        for (i = 0; i < result.warnings.length; i++) {
          html += '<div class="warning-line">' + result.warnings[i] + '</div>';
        }
      }
      html += '<div class="ductstatic-result-note">Straight: Drop = (Length x FR) / 100. Elbow: Drop = (Velocity / 4000)^2 x C.</div>';
      html += '<div class="ductstatic-result-note">Transition and Outlet are included as placeholders until dedicated pressure-drop formulas are provided.</div>';
      return html;
    }

    function calculateDuctStatic() {
      var sel = getDuctStaticSelectionState();
      var warnings = [];
      var total = 0;
      var i;
      var hasMeaningfulInputs = false;
      for (i = 0; i < sel.segments.length; i++) {
        total += parseNonNegativeNumber(sel.segments[i].segmentPressureDrop, 0);
        if (sel.segments[i].warning) warnings.push(sel.segments[i].name + ': ' + sel.segments[i].warning);
        if (sel.segments[i].lineType === 'Straight' && parseNonNegativeNumber(sel.segments[i].lengthFt, 0) > 0 && parseNonNegativeNumber(sel.segments[i].frictionRate, 0) > 0) {
          hasMeaningfulInputs = true;
        }
        if (sel.segments[i].lineType === 'Elbow' &&
            parseNonNegativeNumber(sel.segments[i].velocityFpm, 0) > 0 &&
            parseNonNegativeNumber(sel.segments[i].width, 0) > 0 &&
            parseNonNegativeNumber(sel.segments[i].height, 0) > 0 &&
            parseNonNegativeNumber(sel.segments[i].radius, 0) > 0) {
          hasMeaningfulInputs = true;
        }
      }
      if (sel.segments.length <= 0) {
        latestDuctStaticSnapshot = null;
        var emptyNode = document.getElementById('ductStaticResults');
        setPlaceholderResult(emptyNode, APP_MESSAGES.ductStaticPlaceholder || 'Enter duct segment data to calculate total static pressure drop.');
        return;
      }
      if (!hasMeaningfulInputs) {
        latestDuctStaticSnapshot = null;
        var placeholderNode = document.getElementById('ductStaticResults');
        setPlaceholderResult(placeholderNode, APP_MESSAGES.ductStaticPlaceholder || 'Enter duct segment data to calculate total static pressure drop.');
        return;
      }
      var result = {
        segments: sel.segments,
        totalPressureDrop: roundNumber(total, 5),
        warnings: warnings
      };
      latestDuctStaticSnapshot = { section: 'ductstatic', selection: sel, result: result };
      var node = document.getElementById('ductStaticResults');
      if (node) node.innerHTML = renderDuctStaticResult(result);
    }

    function calculateSystemCharge(chargeInputs) {
      var installed = parseNonNegativeNumber(chargeInputs.installedPipeLength, 0);
      var included = parseNonNegativeNumber(chargeInputs.includedPipeLength, 0);
      var excessLength = Math.max(0, installed - included);
      var calculatedAdditionalCharge = excessLength * parseNonNegativeNumber(chargeInputs.additionalChargeRate, 0);
      var additionalCharge = chargeInputs.useManualAdditionalCharge ? parseNonNegativeNumber(chargeInputs.manualAdditionalCharge, 0) : calculatedAdditionalCharge;
      var ms = parseNonNegativeNumber(chargeInputs.factoryCharge, 0) + parseNonNegativeNumber(chargeInputs.fieldCharge, 0) + additionalCharge;
      return {
        excessLength: excessLength,
        calculatedAdditionalCharge: calculatedAdditionalCharge,
        additionalCharge: additionalCharge,
        ms: ms
      };
    }

    function determineApplicableStandard(classification) {
      if (classification.highProbability !== 'YES') {
        return {
          applicableStandard: '',
          warning: 'This calculator is intended for high-probability A2L HVAC systems. Final compliance requires manual review.',
          reason: 'High-probability is set to No; final pass/fail is not issued.'
        };
      }
      if (classification.spaceType === 'RESIDENTIAL' && classification.oneSystemOneDwelling === 'YES' && classification.highProbability === 'YES') {
        return {
          applicableStandard: 'ASHRAE 15.2',
          warning: '',
          reason: 'ASHRAE 15.2 selected because the system is a one-to-one residential dwelling unit high-probability system.'
        };
      }
      return {
        applicableStandard: 'ASHRAE 15',
        warning: '',
        reason: 'ASHRAE 15 selected because classification is not one-to-one residential, while remaining high-probability.'
      };
    }

    function selectAshrae15Equation(inputs) {
      if (inputs.selectionMode === 'MANUAL') return inputs.manualEquation === 'EQ_7_9' ? 'EQ_7_9' : 'EQ_7_8';
      if (inputs.connectedSpaces === 'YES' && inputs.airCirculation === 'YES') return 'EQ_7_9';
      return 'EQ_7_8';
    }

    function calculateAshrae15Compliance(ms, equation, inputs) {
      var result = {
        methodLabel: equation === 'EQ_7_9' ? 'Using Equation 7-9' : 'Using Equation 7-8',
        equation: equation,
        edvc: 0,
        compliant: false,
        status: 'INCOMPLETE',
        warning: '',
        details: {}
      };
      if (equation === 'EQ_7_8') {
        var roomArea = parseNonNegativeNumber(inputs.roomArea78, -1);
        var ceilingHeight = parseNonNegativeNumber(inputs.ceilingHeight78, -1);
        var lfl = parseNonNegativeNumber(inputs.lfl78, -1);
        var cf = parseNonNegativeNumber(inputs.cf78, -1);
        var focc = parseNonNegativeNumber(inputs.focc78, -1);
        if (roomArea <= 0 || ceilingHeight <= 0 || lfl <= 0 || cf <= 0 || focc <= 0) {
          result.warning = 'Equation 7-8 requires Room Area, Ceiling Height, LFL, CF, and Focc greater than zero.';
          result.status = 'INCOMPLETE';
          return result;
        }
        var veff = roomArea * ceilingHeight;
        var edvc = veff * (lfl / 1000.0) * cf * focc;
        result.edvc = edvc;
        result.compliant = (ms <= edvc);
        result.status = result.compliant ? 'COMPLIANT' : 'NOT COMPLIANT';
        result.details = { veff: veff, lfl: lfl, cf: cf, focc: focc, edvc: edvc };
        return result;
      }

      var mdef = parseNonNegativeNumber(inputs.mdef79, -1);
      var flfl = parseNonNegativeNumber(inputs.flfl79, -1);
      var focc79 = parseNonNegativeNumber(inputs.focc79, -1);
      if (mdef <= 0 || flfl <= 0 || focc79 <= 0) {
        result.warning = 'Equation 7-9 requires Mdef, FLFL, and Focc greater than zero.';
        result.status = 'INCOMPLETE';
        return result;
      }
      var edvc79 = mdef * flfl * focc79;
      result.edvc = edvc79;
      result.compliant = (ms <= edvc79);
      result.status = result.compliant ? 'COMPLIANT' : 'NOT COMPLIANT';
      result.details = {
        mdef: mdef,
        flfl: flfl,
        focc: focc79,
        lowestOpeningHeight: parseNonNegativeNumber(inputs.lowestOpeningHeight79, 0),
        optionalRoomArea: parseNonNegativeNumber(inputs.optionalRoomArea79, 0),
        edvc: edvc79
      };
      return result;
    }

    var REFRIGERANT_C_TABLE_92 = {
      'R-32': 1.00,
      'R-452B': 1.02,
      'R-454A': 0.92,
      'R-454B': 0.97,
      'R-454C': 0.95,
      'R-457A': 0.71
    };
    var REFRIGERANT_M_TABLE_93 = [
      { area: 100, m: 3.4 }, { area: 125, m: 4.3 }, { area: 150, m: 5.2 }, { area: 175, m: 6.0 }, { area: 200, m: 6.9 },
      { area: 225, m: 7.8 }, { area: 250, m: 8.6 }, { area: 275, m: 9.5 }, { area: 300, m: 10.3 }, { area: 325, m: 11.2 },
      { area: 350, m: 12.1 }, { area: 375, m: 12.9 }, { area: 400, m: 13.8 }, { area: 425, m: 14.6 }, { area: 450, m: 15.5 },
      { area: 475, m: 16.4 }, { area: 500, m: 17.2 }, { area: 525, m: 18.1 }, { area: 550, m: 19.0 }, { area: 575, m: 19.8 },
      { area: 600, m: 20.7 }, { area: 625, m: 21.5 }, { area: 650, m: 22.4 }, { area: 675, m: 23.3 }, { area: 700, m: 24.1 },
      { area: 725, m: 25.0 }, { area: 750, m: 25.9 }, { area: 775, m: 26.7 }, { area: 800, m: 27.6 }, { area: 825, m: 28.4 },
      { area: 850, m: 29.3 }, { area: 875, m: 30.2 }, { area: 900, m: 31.0 }, { area: 925, m: 31.9 }, { area: 950, m: 32.7 },
      { area: 975, m: 33.6 }, { area: 1000, m: 34.5 }, { area: 1025, m: 35.1 }
    ];
    var REFRIGERANT_MV_TABLE_94 = [
      { cfm: 20, mv: 0.4 }, { cfm: 40, mv: 0.7 }, { cfm: 60, mv: 1.1 }, { cfm: 80, mv: 1.4 }, { cfm: 100, mv: 1.8 },
      { cfm: 120, mv: 2.1 }, { cfm: 140, mv: 2.5 }, { cfm: 160, mv: 2.8 }, { cfm: 180, mv: 3.2 }, { cfm: 200, mv: 3.5 },
      { cfm: 220, mv: 4.2 }, { cfm: 240, mv: 4.6 }, { cfm: 260, mv: 5.0 }, { cfm: 280, mv: 5.4 }, { cfm: 300, mv: 5.8 },
      { cfm: 320, mv: 6.2 }, { cfm: 340, mv: 6.5 }, { cfm: 360, mv: 6.9 }, { cfm: 380, mv: 7.3 }, { cfm: 400, mv: 7.3 }
    ];

    function linearInterpolate(x, x1, y1, x2, y2) {
      if (Math.abs(x2 - x1) < 1e-9) return y1;
      return y1 + ((x - x1) * (y2 - y1) / (x2 - x1));
    }

    function getCByRefrigerant(refrigerant, useOverride, overrideValue) {
      var value = parseNonNegativeNumber(REFRIGERANT_C_TABLE_92[String(refrigerant || '')], 0);
      if (useOverride) {
        var cOverride = parseNonNegativeNumber(overrideValue, -1);
        if (cOverride > 0) value = cOverride;
      }
      return {
        value: value,
        warning: value > 0 ? '' : 'Selected refrigerant is not available in Table 9-2 C lookup.',
        source: useOverride ? 'MANUAL_OVERRIDE' : 'TABLE_9_2'
      };
    }

    function getMByArea(areaFt2, interpolate, allowExtrapolation) {
      var area = parseNonNegativeNumber(areaFt2, 0);
      var table = REFRIGERANT_M_TABLE_93;
      var first = table[0];
      var last = table[table.length - 1];
      var i;
      if (area <= 0) return { value: 0, warning: 'Area or equivalent area is required for Table 9-3 lookup.' };
      if (area <= first.area) return { value: first.m, warning: area < first.area ? 'Area is below Table 9-3 minimum; using minimum row value.' : '' };
      if (area >= last.area) {
        if (area === last.area) return { value: last.m, warning: '' };
        if (allowExtrapolation) {
          var p1 = table[table.length - 2];
          return { value: linearInterpolate(area, p1.area, p1.m, last.area, last.m), warning: 'Area exceeds Table 9-3 maximum; extrapolated from top range.' };
        }
        return { value: last.m, warning: 'Area exceeds Table 9-3 maximum; capped to highest table row.' };
      }
      for (i = 0; i < table.length - 1; i++) {
        var a = table[i];
        var b = table[i + 1];
        if (area >= a.area && area <= b.area) {
          if (!interpolate) return { value: a.m, warning: '' };
          if (area === a.area) return { value: a.m, warning: '' };
          if (area === b.area) return { value: b.m, warning: '' };
          return { value: linearInterpolate(area, a.area, a.m, b.area, b.m), warning: '' };
        }
      }
      return { value: 0, warning: 'Unable to resolve Table 9-3 lookup.' };
    }

    function getHeightCorrection(heightFt) {
      var h = parseNonNegativeNumber(heightFt, 0);
      if (h <= 0) return { value: 1, warning: 'Space height must be greater than zero.' };
      return { value: Math.min(h / 7.2, 1.25), warning: '' };
    }

    function getMc(areaFt2, heightFt, interpolate, allowExtrapolation) {
      var mLookup = getMByArea(areaFt2, interpolate, allowExtrapolation);
      var hcLookup = getHeightCorrection(heightFt);
      return {
        m: mLookup.value,
        hc: hcLookup.value,
        value: parseNonNegativeNumber(mLookup.value, 0) * parseNonNegativeNumber(hcLookup.value, 1),
        warning: [mLookup.warning, hcLookup.warning].filter(Boolean).join(' | ')
      };
    }

    function getMvByVentilation(ventCfm, interpolate, ventilationUsed) {
      var cfm = parseNonNegativeNumber(ventCfm, 0);
      var table = REFRIGERANT_MV_TABLE_94;
      var i;
      if (!ventilationUsed || cfm <= 0) return { value: 0, warning: '' };
      if (cfm >= 400) return { value: 7.3, warning: '' };
      if (cfm < 20) return { value: 0, warning: 'Ventilation CFM below Table 9-4 minimum; using Mv = 0.' };
      for (i = 0; i < table.length - 1; i++) {
        var a = table[i];
        var b = table[i + 1];
        if (cfm >= a.cfm && cfm <= b.cfm) {
          if (!interpolate) return { value: a.mv, warning: '' };
          if (cfm === a.cfm) return { value: a.mv, warning: '' };
          if (cfm === b.cfm) return { value: b.mv, warning: '' };
          return { value: linearInterpolate(cfm, a.cfm, a.mv, b.cfm, b.mv), warning: '' };
        }
      }
      return { value: 0, warning: 'Unable to resolve Table 9-4 ventilation lookup.' };
    }

    function resolveAreaForA2L(inputs) {
      var roomArea = parseNonNegativeNumber(inputs.roomArea, 0);
      var roomVolume = parseNonNegativeNumber(inputs.roomVolume, 0);
      if (roomArea > 0) return { area: roomArea, source: 'ROOM_AREA' };
      if (roomVolume > 0) return { area: roomVolume / 7.2, source: 'ROOM_VOLUME_EQ_AREA' };
      return { area: 0, source: 'NONE' };
    }

    function calculateA2LRow(rowInput) {
      var warnings = [];
      var useInterpolation = rowInput.interpolate !== false;
      var allowAreaExtrapolation = rowInput.allowAreaExtrapolation === true;
      var numberOfUnits = parseNonNegativeNumber(rowInput.numberOfUnits, 0);
      var lineCharge = parseNonNegativeNumber(rowInput.pipeLengthFt, 0) * parseNonNegativeNumber(rowInput.chargePerFt, 0) * numberOfUnits;
      var totalCharge = parseNonNegativeNumber(rowInput.factoryCharge, 0) + lineCharge;
      var areaInfo = resolveAreaForA2L(rowInput);
      var spaceHeight = parseNonNegativeNumber(rowInput.spaceHeightFt, 0);
      if (spaceHeight <= 0 && parseNonNegativeNumber(rowInput.roomVolume, 0) > 0) spaceHeight = 7.2;
      var cLookup = getCByRefrigerant(rowInput.refrigerantType, rowInput.useCOverride === true, rowInput.cOverride);
      var mcLookup = getMc(areaInfo.area, spaceHeight, useInterpolation, allowAreaExtrapolation);
      var mvLookup = getMvByVentilation(rowInput.ventilationCfm, useInterpolation, rowInput.ventilationUsed === true);
      var m1 = parseNonNegativeNumber(rowInput.m1, 0);
      var m2 = parseNonNegativeNumber(rowInput.m2, 0);
      var compareCharge = parseNonNegativeNumber(rowInput.mrel, 0) > 0 ? parseNonNegativeNumber(rowInput.mrel, 0) : totalCharge;
      var mMaxNoVent = cLookup.value * mcLookup.value;
      var mMaxWithVent = cLookup.value * (mcLookup.value + mvLookup.value);
      var complianceNoVent = 'MANUAL REVIEW';
      var complianceWithVent = 'MANUAL REVIEW';

      if (cLookup.warning) warnings.push(cLookup.warning);
      if (mcLookup.warning) warnings.push(mcLookup.warning);
      if (mvLookup.warning) warnings.push(mvLookup.warning);
      if (m1 > 0 && m2 > 0 && m2 < m1) warnings.push('m2 must be greater than or equal to m1.');

      if (m1 > 0 && compareCharge <= m1) {
        complianceNoVent = 'YES';
        complianceWithVent = 'YES';
      } else if (m2 > 0 && compareCharge > m2) {
        complianceNoVent = 'NO';
        complianceWithVent = 'NO';
      } else {
        if (mMaxNoVent > 0) complianceNoVent = compareCharge <= mMaxNoVent ? 'YES' : 'NO';
        if (mMaxWithVent > 0) complianceWithVent = compareCharge <= mMaxWithVent ? 'YES' : 'NO';
      }

      return {
        lineCharge: lineCharge,
        totalCharge: totalCharge,
        compareCharge: compareCharge,
        c: cLookup.value,
        m: mcLookup.m,
        hc: mcLookup.hc,
        mc: mcLookup.value,
        mv: mvLookup.value,
        mMaxNoVent: mMaxNoVent,
        mMaxWithVent: mMaxWithVent,
        m1: m1,
        m2: m2,
        complianceNoVent: complianceNoVent,
        complianceWithVent: complianceWithVent,
        warnings: warnings,
        areaUsed: areaInfo.area
      };
    }

    function calculateAshrae152Compliance(ms, inputs, advancedInputs, classification) {
      var useMrel = advancedInputs.useMrel === 'YES';
      var mrel = parseNonNegativeNumber(advancedInputs.mrel, 0);
      var row = calculateA2LRow({
        refrigerantType: classification ? classification.refrigerantType : '',
        roomArea: inputs.roomArea,
        roomVolume: inputs.roomVolume,
        spaceHeightFt: inputs.spaceHeight,
        ventilationUsed: inputs.ventilationUsed === 'YES',
        ventilationCfm: inputs.ventilationCfm,
        pipeLengthFt: 0,
        chargePerFt: 0,
        numberOfUnits: 0,
        factoryCharge: ms,
        m1: inputs.m1,
        m2: inputs.m2,
        mrel: useMrel ? mrel : 0,
        interpolate: inputs.interpolateLookups !== false,
        allowAreaExtrapolation: inputs.allowAreaExtrapolation === true,
        useCOverride: inputs.useCOverride === true,
        cOverride: inputs.cOverride
      });
      var compareValue = useMrel ? mrel : ms;
      var compareLabel = useMrel ? 'mrel' : 'ms';
      var methodLabel = '';
      var limitLabel = 'mMAX No Vent';
      var limitValue = row.mMaxNoVent;
      var warning = row.warnings.join(' | ');
      var status = 'INCOMPLETE';
      var compliant = false;

      if (row.m1 <= 0 || row.m2 <= 0) {
        warning = (warning ? warning + ' | ' : '') + 'm1 and m2 are required and must be greater than zero.';
      } else if (compareValue <= row.m1) {
        methodLabel = compareLabel + ' <= m1';
        limitLabel = 'm1';
        limitValue = row.m1;
        status = 'COMPLIANT';
        compliant = true;
      } else if (compareValue > row.m2) {
        methodLabel = compareLabel + ' > m2';
        limitLabel = 'm2';
        limitValue = row.m2;
        status = 'NOT COMPLIANT';
      } else if (row.mMaxNoVent > 0 && compareValue <= row.mMaxNoVent) {
        methodLabel = compareLabel + ' <= mMAX No Vent';
        limitLabel = 'mMAX No Vent';
        limitValue = row.mMaxNoVent;
        status = 'COMPLIANT';
        compliant = true;
      } else if (row.mMaxWithVent > 0 && compareValue <= row.mMaxWithVent) {
        methodLabel = compareLabel + ' <= mMAX With Vent';
        limitLabel = 'mMAX With Vent';
        limitValue = row.mMaxWithVent;
        status = 'COMPLIANT';
        compliant = true;
      } else if (row.mMaxWithVent > 0) {
        methodLabel = compareLabel + ' > mMAX With Vent';
        limitLabel = 'mMAX With Vent';
        limitValue = row.mMaxWithVent;
        status = 'NOT COMPLIANT';
      }

      return {
        methodLabel: methodLabel,
        compliant: compliant,
        status: status,
        warning: warning,
        limitLabel: limitLabel,
        limitValue: limitValue,
        m1: row.m1,
        m2: row.m2,
        mmax: row.mMaxNoVent,
        compareValue: compareValue,
        compareLabel: compareLabel,
        details: {
          c: row.c,
          m: row.m,
          hc: row.hc,
          mc: row.mc,
          mv: row.mv,
          mmaxNoVent: row.mMaxNoVent,
          mmaxWithVent: row.mMaxWithVent,
          roomArea: parseNonNegativeNumber(inputs.roomArea, 0),
          roomVolume: parseNonNegativeNumber(inputs.roomVolume, 0),
          spaceHeight: parseNonNegativeNumber(inputs.spaceHeight, 0),
          ventilationCfm: parseNonNegativeNumber(inputs.ventilationCfm, 0),
          mrel: useMrel ? mrel : 0,
          statement: status === 'COMPLIANT' ? 'System charge is within allowable limit.' : (status === 'NOT COMPLIANT' ? 'System charge exceeds allowable limit.' : 'Enter required inputs to evaluate compliance.')
        }
      };
    }

    function makeRefrigerantMultiRowId() {
      refrigerantMultiRowIdSeq += 1;
      return 'refrig_row_' + refrigerantMultiRowIdSeq;
    }

    function getRefrigerantMultiSharedState() {
      return {
        refrigerantType: (document.getElementById('refMzRefrigerantType') ? document.getElementById('refMzRefrigerantType').value : 'R-454B'),
        safetyGroup: 'A2L',
        standardPath: (document.getElementById('refMzStandardPath') ? document.getElementById('refMzStandardPath').value : 'AUTO'),
        dataVersion: (document.getElementById('refMzDataVersion') ? document.getElementById('refMzDataVersion').value : 'ADDENDA'),
        defaultPipeSize: (document.getElementById('refMzDefaultPipeSize') ? String(document.getElementById('refMzDefaultPipeSize').value || '').trim() : ''),
        defaultChargePerFt: parseNonNegativeNumber(document.getElementById('refMzDefaultChargePerFt') ? document.getElementById('refMzDefaultChargePerFt').value : 0, 0),
        defaultFactoryCharge: parseNonNegativeNumber(document.getElementById('refMzDefaultFactoryCharge') ? document.getElementById('refMzDefaultFactoryCharge').value : 0, 0),
        ventilationConsidered: (document.getElementById('refMzVentilationConsidered') ? document.getElementById('refMzVentilationConsidered').value : 'NO'),
        ssovConsidered: (document.getElementById('refMzSsovConsidered') ? document.getElementById('refMzSsovConsidered').value : 'NO'),
        defaultUnits: parseNonNegativeNumber(document.getElementById('refMzDefaultUnits') ? document.getElementById('refMzDefaultUnits').value : 1, 1),
        defaultSpaceHeight: parseNonNegativeNumber(document.getElementById('refMzDefaultSpaceHeight') ? document.getElementById('refMzDefaultSpaceHeight').value : 7.2, 7.2),
        defaultVentCfm: parseNonNegativeNumber(document.getElementById('refMzDefaultVentCfm') ? document.getElementById('refMzDefaultVentCfm').value : 0, 0),
        defaultM1: parseNonNegativeNumber(document.getElementById('refMzDefaultM1') ? document.getElementById('refMzDefaultM1').value : 0, 0),
        defaultM2: parseNonNegativeNumber(document.getElementById('refMzDefaultM2') ? document.getElementById('refMzDefaultM2').value : 0, 0),
        interpolateLookups: (document.getElementById('refMzInterpolate') ? document.getElementById('refMzInterpolate').value : 'YES') === 'YES',
        allowAreaExtrapolation: (document.getElementById('refMzAllowAreaExtrapolation') ? document.getElementById('refMzAllowAreaExtrapolation').value : 'NO') === 'YES'
      };
    }

    function getRefrigerantSelectedMultiRowId() {
      var body = document.getElementById('refMultiRowsBody');
      if (!body) return '';
      var checks = body.getElementsByClassName('ref-mz-select');
      var i;
      for (i = 0; i < checks.length; i++) {
        if (checks[i].checked) return checks[i].getAttribute('data-row-id') || '';
      }
      return '';
    }

    function applyRefrigerantSharedDefaultsToRow(row, shared) {
      row.pipeSize = row.pipeSize || shared.defaultPipeSize || '3/8 in';
      row.numberOfUnits = parseNonNegativeNumber((row.numberOfUnits === undefined || row.numberOfUnits === null) ? shared.defaultUnits : row.numberOfUnits, shared.defaultUnits);
      row.chargePerFoot = parseNonNegativeNumber((row.chargePerFoot === undefined || row.chargePerFoot === null) ? shared.defaultChargePerFt : row.chargePerFoot, shared.defaultChargePerFt);
      row.factoryCharge = parseNonNegativeNumber((row.factoryCharge === undefined || row.factoryCharge === null) ? shared.defaultFactoryCharge : row.factoryCharge, shared.defaultFactoryCharge);
      row.spaceHeight = parseNonNegativeNumber((row.spaceHeight === undefined || row.spaceHeight === null) ? ((row.dispersalHeight === undefined || row.dispersalHeight === null) ? shared.defaultSpaceHeight : row.dispersalHeight) : row.spaceHeight, shared.defaultSpaceHeight);
      row.roomVolume = parseNonNegativeNumber((row.roomVolume === undefined || row.roomVolume === null) ? 0 : row.roomVolume, 0);
      row.roomArea = parseNonNegativeNumber((row.roomArea === undefined || row.roomArea === null) ? 0 : row.roomArea, 0);
      row.pipeLength = parseNonNegativeNumber((row.pipeLength === undefined || row.pipeLength === null) ? 0 : row.pipeLength, 0);
      row.ventilationCfm = parseNonNegativeNumber((row.ventilationCfm === undefined || row.ventilationCfm === null) ? shared.defaultVentCfm : row.ventilationCfm, shared.defaultVentCfm);
      row.mrel = parseNonNegativeNumber((row.mrel === undefined || row.mrel === null) ? 0 : row.mrel, 0);
      row.m1 = parseNonNegativeNumber((row.m1 === undefined || row.m1 === null) ? shared.defaultM1 : row.m1, shared.defaultM1);
      row.m2 = parseNonNegativeNumber((row.m2 === undefined || row.m2 === null) ? shared.defaultM2 : row.m2, shared.defaultM2);
      return row;
    }

    function addRefrigerantMultiRow(seedRow) {
      var shared = getRefrigerantMultiSharedState();
      var row = seedRow && typeof seedRow === 'object' ? JSON.parse(JSON.stringify(seedRow)) : {};
      row.id = row.id || makeRefrigerantMultiRowId();
      var m = /^refrig_row_(\d+)$/.exec(String(row.id || ''));
      if (m && parseInt(m[1], 10) > refrigerantMultiRowIdSeq) refrigerantMultiRowIdSeq = parseInt(m[1], 10);
      row.roomId = row.roomId || ('Room ' + (refrigerantMultiRows.length + 1));
      row = applyRefrigerantSharedDefaultsToRow(row, shared);
      refrigerantMultiRows.push(row);
      renderRefrigerantMultiRows();
      calculateRefrigerantCompliance(false);
    }

    function clearRefrigerantMultiRows() {
      refrigerantMultiRows = [];
      addRefrigerantMultiRow();
    }

    function duplicateRefrigerantSelectedRow() {
      var selectedId = getRefrigerantSelectedMultiRowId();
      var i, source;
      if (!selectedId) {
        alert('Select a row to duplicate.');
        return;
      }
      for (i = 0; i < refrigerantMultiRows.length; i++) {
        if (refrigerantMultiRows[i].id === selectedId) {
          source = JSON.parse(JSON.stringify(refrigerantMultiRows[i]));
          source.id = makeRefrigerantMultiRowId();
          source.roomId = (source.roomId || 'Room') + ' Copy';
          refrigerantMultiRows.splice(i + 1, 0, source);
          break;
        }
      }
      renderRefrigerantMultiRows();
      calculateRefrigerantCompliance(false);
    }

    function deleteRefrigerantSelectedRow() {
      var selectedId = getRefrigerantSelectedMultiRowId();
      var i;
      if (!selectedId) {
        alert('Select a row to delete.');
        return;
      }
      if (refrigerantMultiRows.length <= 1) {
        alert('At least one row is required.');
        return;
      }
      for (i = refrigerantMultiRows.length - 1; i >= 0; i--) {
        if (refrigerantMultiRows[i].id === selectedId) refrigerantMultiRows.splice(i, 1);
      }
      renderRefrigerantMultiRows();
      calculateRefrigerantCompliance(false);
    }

    function removeRefrigerantMultiRowById(rowId) {
      var i;
      if (refrigerantMultiRows.length <= 1) {
        alert('At least one row is required.');
        return;
      }
      for (i = refrigerantMultiRows.length - 1; i >= 0; i--) {
        if (refrigerantMultiRows[i].id === rowId) refrigerantMultiRows.splice(i, 1);
      }
      renderRefrigerantMultiRows();
      calculateRefrigerantCompliance(false);
    }

    function updateRefrigerantMultiRowField(rowId, fieldName, rawValue) {
      var i, row;
      for (i = 0; i < refrigerantMultiRows.length; i++) {
        row = refrigerantMultiRows[i];
        if (row.id !== rowId) continue;
        if (fieldName === 'roomId' || fieldName === 'pipeSize') row[fieldName] = String(rawValue || '');
        else row[fieldName] = parseNonNegativeNumber(rawValue, 0);
        break;
      }
      calculateRefrigerantCompliance(false);
    }

    function applyRefrigerantSharedDefaultsToAllRows() {
      var shared = getRefrigerantMultiSharedState();
      var i;
      for (i = 0; i < refrigerantMultiRows.length; i++) {
        refrigerantMultiRows[i].pipeSize = shared.defaultPipeSize || refrigerantMultiRows[i].pipeSize;
        refrigerantMultiRows[i].numberOfUnits = shared.defaultUnits;
        refrigerantMultiRows[i].chargePerFoot = shared.defaultChargePerFt;
        refrigerantMultiRows[i].factoryCharge = shared.defaultFactoryCharge;
        refrigerantMultiRows[i].spaceHeight = shared.defaultSpaceHeight;
        refrigerantMultiRows[i].ventilationCfm = shared.defaultVentCfm;
        refrigerantMultiRows[i].m1 = shared.defaultM1;
        refrigerantMultiRows[i].m2 = shared.defaultM2;
      }
      renderRefrigerantMultiRows();
      calculateRefrigerantCompliance(false);
    }

    function renderRefrigerantMultiRows() {
      var body = document.getElementById('refMultiRowsBody');
      var html = '';
      var i, row;
      if (!body) return;
      for (i = 0; i < refrigerantMultiRows.length; i++) {
        row = refrigerantMultiRows[i];
        html += '<tr>';
        html += '<td class="center"><input type="radio" name="refMzSelectedRow" class="ref-mz-select" data-row-id="' + row.id + '" /></td>';
        html += '<td><input type="text" value="' + (row.roomId || '') + '" oninput="updateRefrigerantMultiRowField(\'' + row.id + '\',\'roomId\',this.value);" /></td>';
        html += '<td class="num"><input type="number" min="0" step="0.01" value="' + parseNonNegativeNumber(row.roomVolume, 0) + '" oninput="updateRefrigerantMultiRowField(\'' + row.id + '\',\'roomVolume\',this.value);" /></td>';
        html += '<td class="num"><input type="number" min="0" step="0.01" value="' + parseNonNegativeNumber(row.roomArea, 0) + '" oninput="updateRefrigerantMultiRowField(\'' + row.id + '\',\'roomArea\',this.value);" /></td>';
        html += '<td class="num"><input type="number" min="0" step="0.01" value="' + parseNonNegativeNumber(row.spaceHeight, 7.2) + '" oninput="updateRefrigerantMultiRowField(\'' + row.id + '\',\'spaceHeight\',this.value);" /></td>';
        html += '<td class="num"><input type="number" min="0" step="0.01" value="' + parseNonNegativeNumber(row.pipeLength, 0) + '" oninput="updateRefrigerantMultiRowField(\'' + row.id + '\',\'pipeLength\',this.value);" /></td>';
        html += '<td><input type="text" value="' + (row.pipeSize || '') + '" oninput="updateRefrigerantMultiRowField(\'' + row.id + '\',\'pipeSize\',this.value);" /></td>';
        html += '<td class="num"><input type="number" min="0" step="1" value="' + parseNonNegativeNumber(row.numberOfUnits, 0) + '" oninput="updateRefrigerantMultiRowField(\'' + row.id + '\',\'numberOfUnits\',this.value);" /></td>';
        html += '<td class="num"><input type="number" min="0" step="0.0001" value="' + parseNonNegativeNumber(row.chargePerFoot, 0) + '" oninput="updateRefrigerantMultiRowField(\'' + row.id + '\',\'chargePerFoot\',this.value);" /></td>';
        html += '<td class="num"><input type="number" min="0" step="0.01" value="' + parseNonNegativeNumber(row.factoryCharge, 0) + '" oninput="updateRefrigerantMultiRowField(\'' + row.id + '\',\'factoryCharge\',this.value);" /></td>';
        html += '<td class="num readonly">-</td>';
        html += '<td class="num readonly">-</td>';
        html += '<td class="num"><input type="number" min="0" step="0.01" value="' + parseNonNegativeNumber(row.mrel, 0) + '" oninput="updateRefrigerantMultiRowField(\'' + row.id + '\',\'mrel\',this.value);" /></td>';
        html += '<td class="num"><input type="number" min="0" step="0.01" value="' + parseNonNegativeNumber(row.m1, 0) + '" oninput="updateRefrigerantMultiRowField(\'' + row.id + '\',\'m1\',this.value);" /></td>';
        html += '<td class="num"><input type="number" min="0" step="0.01" value="' + parseNonNegativeNumber(row.m2, 0) + '" oninput="updateRefrigerantMultiRowField(\'' + row.id + '\',\'m2\',this.value);" /></td>';
        html += '<td class="num readonly">-</td>';
        html += '<td class="num readonly">-</td>';
        html += '<td class="num readonly">-</td>';
        html += '<td class="num readonly">-</td>';
        html += '<td class="num readonly">-</td>';
        html += '<td class="num readonly">-</td>';
        html += '<td class="num readonly">-</td>';
        html += '<td class="readonly">-</td>';
        html += '<td class="readonly">-</td>';
        html += '<td><input type="button" class="btn btn-small" value="Remove" onclick="removeRefrigerantMultiRowById(\'' + row.id + '\');" /></td>';
        html += '</tr>';
      }
      body.innerHTML = html;
    }

    function evaluateRefrigerantMultiRow(row, shared, classification) {
      var calc = calculateA2LRow({
        refrigerantType: shared.refrigerantType,
        roomArea: row.roomArea,
        roomVolume: row.roomVolume,
        spaceHeightFt: row.spaceHeight,
        ventilationUsed: shared.ventilationConsidered === 'YES',
        ventilationCfm: row.ventilationCfm,
        pipeLengthFt: row.pipeLength,
        chargePerFt: row.chargePerFoot,
        numberOfUnits: row.numberOfUnits,
        factoryCharge: row.factoryCharge,
        m1: row.m1,
        m2: row.m2,
        mrel: row.mrel,
        interpolate: shared.interpolateLookups !== false,
        allowAreaExtrapolation: shared.allowAreaExtrapolation === true
      });
      var statusNoVent = calc.complianceNoVent === 'YES' ? 'COMPLIANT' : (calc.complianceNoVent === 'NO' ? 'NOT COMPLIANT' : 'MANUAL REVIEW REQUIRED');
      var statusWithVent = calc.complianceWithVent === 'YES' ? 'COMPLIANT' : (calc.complianceWithVent === 'NO' ? 'NOT COMPLIANT' : 'MANUAL REVIEW REQUIRED');
      if (shared.standardPath === 'ASHRAE_15') {
        statusNoVent = 'MANUAL REVIEW REQUIRED';
        statusWithVent = 'MANUAL REVIEW REQUIRED';
      }
      return {
        rowId: row.id,
        roomId: row.roomId || '',
        roomArea: parseNonNegativeNumber(row.roomArea, 0),
        roomVolume: parseNonNegativeNumber(row.roomVolume, 0),
        spaceHeight: parseNonNegativeNumber(row.spaceHeight, 0),
        pipeLength: parseNonNegativeNumber(row.pipeLength, 0),
        pipeSize: row.pipeSize || '',
        numberOfUnits: parseNonNegativeNumber(row.numberOfUnits, 0),
        chargePerFoot: parseNonNegativeNumber(row.chargePerFoot, 0),
        factoryCharge: parseNonNegativeNumber(row.factoryCharge, 0),
        lineCharge: calc.lineCharge,
        totalCharge: calc.totalCharge,
        mrel: parseNonNegativeNumber(row.mrel, 0),
        m1: calc.m1,
        m2: calc.m2,
        c: calc.c,
        m: calc.m,
        hc: calc.hc,
        mc: calc.mc,
        mv: calc.mv,
        mmaxNoVent: calc.mMaxNoVent,
        mmaxWithVent: calc.mMaxWithVent,
        statusNoVent: statusNoVent,
        statusWithVent: statusWithVent,
        complianceNoVent: calc.complianceNoVent,
        complianceWithVent: calc.complianceWithVent,
        compareCharge: calc.compareCharge,
        selectedStandard: shared.standardPath === 'AUTO' ? determineApplicableStandard(classification).applicableStandard : (shared.standardPath === 'ASHRAE_15' ? 'ASHRAE 15' : 'ASHRAE 15.2'),
        warning: calc.warnings.length ? calc.warnings.join(' | ') : ''
      };
    }

    function calculateRefrigerantMultiZone(shared, classification) {
      var rowsOut = [];
      var i;
      var compliant = 0;
      var notCompliant = 0;
      var manual = 0;
      var warnings = [];
      for (i = 0; i < refrigerantMultiRows.length; i++) {
        var rowResult = evaluateRefrigerantMultiRow(refrigerantMultiRows[i], shared, classification);
        rowsOut.push(rowResult);
        if (rowResult.warning) warnings.push((rowResult.roomId || ('Row ' + (i + 1))) + ': ' + rowResult.warning);
        if (rowResult.statusNoVent === 'COMPLIANT' && rowResult.statusWithVent === 'COMPLIANT') compliant += 1;
        else if (rowResult.statusNoVent === 'NOT COMPLIANT' || rowResult.statusWithVent === 'NOT COMPLIANT') notCompliant += 1;
        else manual += 1;
      }
      var finalStatus = 'COMPLIANT';
      if (notCompliant > 0) finalStatus = 'NOT COMPLIANT';
      else if (manual > 0) finalStatus = 'MANUAL REVIEW REQUIRED';
      return {
        mode: 'MULTI_ZONE',
        section: 'refrigerant',
        timestamp: new Date().toISOString(),
        applicableStandard: (shared.standardPath === 'AUTO') ? determineApplicableStandard(classification).applicableStandard : (shared.standardPath === 'ASHRAE_15' ? 'ASHRAE 15' : 'ASHRAE 15.2'),
        status: finalStatus,
        warning: ((shared.standardPath === 'ASHRAE_15' ? 'Multi-Zone Table is optimized for ASHRAE 15.2 repeated calculations; ASHRAE 15 rows are flagged for manual review.' : '') + (warnings.length ? ((shared.standardPath === 'ASHRAE_15' ? ' ' : '') + warnings.join(' | ')) : '')).trim(),
        shared: shared,
        rows: rowsOut,
        summary: {
          totalRows: rowsOut.length,
          compliantRows: compliant,
          nonCompliantRows: notCompliant,
          manualReviewRows: manual
        }
      };
    }

    function buildRefrigerantComplianceResult(state) {
      var standardInfo = determineApplicableStandard(state.classification);
      var charge = calculateSystemCharge(state.chargeInputs);
      var base = {
        section: 'refrigerant',
        mode: 'SINGLE',
        applicableStandard: standardInfo.applicableStandard,
        standardReason: standardInfo.reason || '',
        classificationWarning: standardInfo.warning,
        refrigerantType: state.classification.refrigerantType,
        safetyGroup: 'A2L',
        systemType: state.classification.systemType,
        spaceType: state.classification.spaceType,
        dataVersion: state.classification.dataVersion,
        charge: charge,
        methodLabel: '',
        limitValue: 0,
        limitLabel: '',
        status: 'MANUAL REVIEW REQUIRED',
        compliant: false,
        warning: '',
        details: {}
      };

      if (standardInfo.warning) {
        base.warning = standardInfo.warning;
        return base;
      }

      if (standardInfo.applicableStandard === 'ASHRAE 15') {
        var equation = selectAshrae15Equation(state.ashrae15Inputs);
        var ash15 = calculateAshrae15Compliance(charge.ms, equation, state.ashrae15Inputs);
        base.methodLabel = ash15.methodLabel;
        base.limitValue = ash15.edvc;
        base.limitLabel = 'EDVC';
        base.status = ash15.status;
        base.compliant = ash15.compliant;
        base.warning = ash15.warning || '';
        base.details = ash15.details || {};
        return base;
      }

      var ash152 = calculateAshrae152Compliance(charge.ms, state.ashrae152Inputs, state.advancedInputs, state.classification);
      base.methodLabel = ash152.methodLabel;
      base.limitValue = ash152.limitValue || 0;
      base.limitLabel = ash152.limitLabel || 'mmax';
      base.status = ash152.status;
      base.compliant = ash152.compliant;
      base.warning = ash152.warning || '';
      base.details = ash152.details || {};
      base.details.m1 = ash152.m1;
      base.details.m2 = ash152.m2;
      base.details.mmax = ash152.mmax;
      base.details.compareValue = ash152.compareValue;
      base.details.compareLabel = ash152.compareLabel;
      return base;
    }

    function getRefrigerantState() {
      var calcMode = (document.getElementById('refCalculationMode') ? document.getElementById('refCalculationMode').value : 'SINGLE');
      return {
        calculationMode: calcMode,
        classification: {
          refrigerantType: (document.getElementById('refRefrigerantType') ? document.getElementById('refRefrigerantType').value : 'R-454B'),
          systemType: (document.getElementById('refSystemType') ? document.getElementById('refSystemType').value : 'SPLIT_DX'),
          spaceType: (document.getElementById('refSpaceType') ? document.getElementById('refSpaceType').value : 'RESIDENTIAL'),
          oneSystemOneDwelling: (document.getElementById('refOneSystemOneDwelling') ? document.getElementById('refOneSystemOneDwelling').value : 'YES'),
          highProbability: (document.getElementById('refHighProbability') ? document.getElementById('refHighProbability').value : 'YES'),
          dataVersion: (document.getElementById('refDataVersion') ? document.getElementById('refDataVersion').value : 'ADDENDA')
        },
        chargeInputs: {
          factoryCharge: parseNonNegativeNumber(document.getElementById('refFactoryCharge') ? document.getElementById('refFactoryCharge').value : 0, 0),
          fieldCharge: parseNonNegativeNumber(document.getElementById('refFieldCharge') ? document.getElementById('refFieldCharge').value : 0, 0),
          installedPipeLength: parseNonNegativeNumber(document.getElementById('refInstalledPipeLength') ? document.getElementById('refInstalledPipeLength').value : 0, 0),
          includedPipeLength: parseNonNegativeNumber(document.getElementById('refIncludedPipeLength') ? document.getElementById('refIncludedPipeLength').value : 0, 0),
          additionalChargeRate: parseNonNegativeNumber(document.getElementById('refAdditionalChargeRate') ? document.getElementById('refAdditionalChargeRate').value : 0, 0),
          useManualAdditionalCharge: !!(document.getElementById('refUseManualAdditionalCharge') && document.getElementById('refUseManualAdditionalCharge').checked),
          manualAdditionalCharge: parseNonNegativeNumber(document.getElementById('refManualAdditionalCharge') ? document.getElementById('refManualAdditionalCharge').value : 0, 0)
        },
        ashrae15Inputs: {
          selectionMode: (document.getElementById('refEqSelectionMode') ? document.getElementById('refEqSelectionMode').value : 'AUTO'),
          connectedSpaces: (document.getElementById('refConnectedSpaces') ? document.getElementById('refConnectedSpaces').value : 'NO'),
          airCirculation: (document.getElementById('refAirCirculation') ? document.getElementById('refAirCirculation').value : 'YES'),
          manualEquation: (document.getElementById('refManualEquation') ? document.getElementById('refManualEquation').value : 'EQ_7_8'),
          roomArea78: parseNonNegativeNumber(document.getElementById('refRoomArea78') ? document.getElementById('refRoomArea78').value : 0, 0),
          ceilingHeight78: parseNonNegativeNumber(document.getElementById('refCeilingHeight78') ? document.getElementById('refCeilingHeight78').value : 0, 0),
          lfl78: parseNonNegativeNumber(document.getElementById('refLfl78') ? document.getElementById('refLfl78').value : 0, 0),
          cf78: parseNonNegativeNumber(document.getElementById('refCf78') ? document.getElementById('refCf78').value : 0, 0),
          focc78: parseNonNegativeNumber(document.getElementById('refFocc78') ? document.getElementById('refFocc78').value : 0, 0),
          mdef79: parseNonNegativeNumber(document.getElementById('refMdef79') ? document.getElementById('refMdef79').value : 0, 0),
          flfl79: parseNonNegativeNumber(document.getElementById('refFlfl79') ? document.getElementById('refFlfl79').value : 0, 0),
          focc79: parseNonNegativeNumber(document.getElementById('refFocc79') ? document.getElementById('refFocc79').value : 0, 0),
          lowestOpeningHeight79: parseNonNegativeNumber(document.getElementById('refLowestOpeningHeight79') ? document.getElementById('refLowestOpeningHeight79').value : 0, 0),
          optionalRoomArea79: parseNonNegativeNumber(document.getElementById('refOptionalRoomArea79') ? document.getElementById('refOptionalRoomArea79').value : 0, 0)
        },
        ashrae152Inputs: {
          m1: parseNonNegativeNumber(document.getElementById('refM1_152') ? document.getElementById('refM1_152').value : 0, 0),
          m2: parseNonNegativeNumber(document.getElementById('refM2_152') ? document.getElementById('refM2_152').value : 0, 0),
          roomArea: parseNonNegativeNumber(document.getElementById('refRoomArea152') ? document.getElementById('refRoomArea152').value : 0, 0),
          roomVolume: parseNonNegativeNumber(document.getElementById('refRoomVolume152') ? document.getElementById('refRoomVolume152').value : 0, 0),
          spaceHeight: parseNonNegativeNumber(document.getElementById('refSpaceHeight152') ? document.getElementById('refSpaceHeight152').value : 7.2, 7.2),
          ventilationCfm: parseNonNegativeNumber(document.getElementById('refVentilationCfm152') ? document.getElementById('refVentilationCfm152').value : 0, 0),
          interpolateLookups: (document.getElementById('refInterpolate152') ? document.getElementById('refInterpolate152').value : 'YES') === 'YES',
          allowAreaExtrapolation: (document.getElementById('refAllowAreaExtrapolation152') ? document.getElementById('refAllowAreaExtrapolation152').value : 'NO') === 'YES',
          useCOverride: (document.getElementById('refUseCOverride152') ? document.getElementById('refUseCOverride152').value : 'NO') === 'YES',
          cOverride: parseNonNegativeNumber(document.getElementById('refCOverride152') ? document.getElementById('refCOverride152').value : 0, 0),
          ventilationUsed: (document.getElementById('refVentilation152') ? document.getElementById('refVentilation152').value : 'NO'),
          ssovUsed: (document.getElementById('refSsov152') ? document.getElementById('refSsov152').value : 'NO')
        },
        advancedInputs: {
          useMrel: (document.getElementById('refUseMrel152') ? document.getElementById('refUseMrel152').value : 'NO'),
          mrel: parseNonNegativeNumber(document.getElementById('refMrel152') ? document.getElementById('refMrel152').value : 0, 0)
        },
        multiZone: {
          shared: getRefrigerantMultiSharedState(),
          rows: refrigerantMultiRows.slice()
        }
      };
    }

    function renderRefrigerantResults(result) {
      var statusClass = 'review';
      if (result.status === 'COMPLIANT') statusClass = 'ok';
      if (result.status === 'NOT COMPLIANT') statusClass = 'fail';
      if (result.status === 'INCOMPLETE') statusClass = 'incomplete';
      var limitText = result.limitLabel ? (result.limitLabel + ': ' + formatNumber(parseNonNegativeNumber(result.limitValue, 0), 2) + ' lb') : '-';
      var html = '';
      if (result.mode === 'MULTI_ZONE') {
        var summary = result.summary || {};
        var rows = result.rows || [];
        var i, rowStatus;
        html += '<div class="refrig-result-main">';
        html += '<div class="result-line"><strong>Applicable Standard:</strong> ' + (result.applicableStandard || '-') + '</div>';
        html += '<div class="result-line"><strong>Calculated At:</strong> ' + (result.timestamp || '-') + '</div>';
        html += '<div class="result-line"><strong>Total Rows:</strong> <span class="value">' + parseNonNegativeNumber(summary.totalRows, 0) + '</span></div>';
        html += '<div class="result-line"><strong>Compliant Rows:</strong> ' + parseNonNegativeNumber(summary.compliantRows, 0) + '</div>';
        html += '<div class="result-line"><strong>Non-Compliant Rows:</strong> ' + parseNonNegativeNumber(summary.nonCompliantRows, 0) + '</div>';
        html += '<div class="result-line"><strong>Manual Review Rows:</strong> ' + parseNonNegativeNumber(summary.manualReviewRows, 0) + '</div>';
        html += '</div>';
        html += '<table class="refrig-table"><thead><tr><th>Room ID</th><th class="num">Governing Charge (lb)</th><th class="num">No Vent Limit (lb)</th><th class="num">With Vent Limit (lb)</th><th>Status (No Vent)</th><th>Status (With Vent)</th></tr></thead><tbody>';
        for (i = 0; i < rows.length; i++) {
          rowStatus = rows[i];
          html += '<tr>';
          html += '<td>' + (rowStatus.roomId || ('Row ' + (i + 1))) + '</td>';
          html += '<td class="num">' + formatNumber(parseNonNegativeNumber(rowStatus.compareCharge, 0), 2) + '</td>';
          html += '<td class="num">' + formatNumber(parseNonNegativeNumber(rowStatus.mmaxNoVent, 0), 2) + '</td>';
          html += '<td class="num">' + formatNumber(parseNonNegativeNumber(rowStatus.mmaxWithVent, 0), 2) + '</td>';
          html += '<td><span class="refrig-status ' + (rowStatus.statusNoVent === 'COMPLIANT' ? 'ok' : (rowStatus.statusNoVent === 'NOT COMPLIANT' ? 'fail' : 'review')) + '">' + rowStatus.statusNoVent + '</span></td>';
          html += '<td><span class="refrig-status ' + (rowStatus.statusWithVent === 'COMPLIANT' ? 'ok' : (rowStatus.statusWithVent === 'NOT COMPLIANT' ? 'fail' : 'review')) + '">' + rowStatus.statusWithVent + '</span></td>';
          html += '</tr>';
        }
        html += '</tbody></table>';
        if (result.warning) html += '<div class="warning-line">' + result.warning + '</div>';
        var multiConclusion = 'Additional mitigation/manual review required.';
        if (result.status === 'COMPLIANT') multiConclusion = 'All rows are within allowable limits.';
        if (result.status === 'NOT COMPLIANT') multiConclusion = 'One or more rows exceed allowable limits.';
        html += '<div class="module-note"><strong>' + multiConclusion + '</strong></div>';
        return html;
      }
      html += '<div class="refrig-result-main refrig-badges">';
      html += '<span class="refrig-live-pill"><strong>Standard:</strong> ' + (result.applicableStandard || '-') + '</span>';
      html += '<span class="refrig-live-pill"><strong>Method:</strong> ' + (result.methodLabel || '-') + '</span>';
      html += '<span class="refrig-live-pill"><strong>Refrigerant:</strong> ' + (result.refrigerantType || '-') + '</span>';
      html += '<span class="refrig-live-pill status-pill"><strong>Status:</strong> <span class="refrig-status ' + statusClass + '">' + (result.status || 'INCOMPLETE') + '</span></span>';
      html += '</div>';

      var marginValue = parseNonNegativeNumber(result.limitValue, 0) - parseNonNegativeNumber(result.charge.ms, 0);
      html += '<div class="refrig-result-grid">';
      html += '<div class="refrig-metric"><span class="module-note">Total Charge</span><div class="value">' + formatNumber(parseNonNegativeNumber(result.charge.ms, 0), 2) + ' lb</div></div>';
      html += '<div class="refrig-metric"><span class="module-note">Limit Value</span><div class="value">' + limitText + '</div></div>';
      html += '<div class="refrig-metric"><span class="module-note">Margin</span><div class="value">' + formatNumber(marginValue, 2) + ' lb</div></div>';
      html += '</div>';
      if (result.status === 'INCOMPLETE') html += '<div class="refrig-margin-note">Enter required inputs to evaluate compliance.</div>';
      else if (marginValue >= 0) html += '<div class="refrig-margin-note">Below limit by ' + formatNumber(marginValue, 2) + ' lb.</div>';
      else html += '<div class="refrig-margin-note">Exceeds limit by ' + formatNumber(Math.abs(marginValue), 2) + ' lb.</div>';

      if (result.applicableStandard === 'ASHRAE 15') {
        html += '<table class="refrig-table"><thead><tr><th>Parameter</th><th class="num">Value</th></tr></thead><tbody>';
        if (result.details.veff !== undefined) {
          html += '<tr><td>Veff</td><td class="num">' + formatNumber(parseNonNegativeNumber(result.details.veff, 0), 2) + ' ft3</td></tr>';
          html += '<tr><td>LFL</td><td class="num">' + formatNumber(parseNonNegativeNumber(result.details.lfl, 0), 2) + ' lb/1000 ft3</td></tr>';
          html += '<tr><td>CF</td><td class="num">' + formatNumber(parseNonNegativeNumber(result.details.cf, 0), 2) + '</td></tr>';
          html += '<tr><td>Focc</td><td class="num">' + formatNumber(parseNonNegativeNumber(result.details.focc, 0), 2) + '</td></tr>';
          html += '<tr><td>EDVC</td><td class="num">' + formatNumber(parseNonNegativeNumber(result.details.edvc, 0), 2) + ' lb</td></tr>';
        } else {
          html += '<tr><td>Mdef</td><td class="num">' + formatNumber(parseNonNegativeNumber(result.details.mdef, 0), 2) + ' lb</td></tr>';
          html += '<tr><td>FLFL</td><td class="num">' + formatNumber(parseNonNegativeNumber(result.details.flfl, 0), 2) + '</td></tr>';
          html += '<tr><td>Focc</td><td class="num">' + formatNumber(parseNonNegativeNumber(result.details.focc, 0), 2) + '</td></tr>';
          html += '<tr><td>EDVC</td><td class="num">' + formatNumber(parseNonNegativeNumber(result.details.edvc, 0), 2) + ' lb</td></tr>';
        }
        html += '</tbody></table>';
      } else if (result.applicableStandard === 'ASHRAE 15.2') {
        html += '<table class="refrig-table"><thead><tr><th>Parameter</th><th class="num">Value</th></tr></thead><tbody>';
        html += '<tr><td>m1</td><td class="num">' + formatNumber(parseNonNegativeNumber(result.details.m1, 0), 2) + ' lb</td></tr>';
        html += '<tr><td>m2</td><td class="num">' + formatNumber(parseNonNegativeNumber(result.details.m2, 0), 2) + ' lb</td></tr>';
        html += '<tr><td>C (Table 9-2)</td><td class="num">' + formatNumber(parseNonNegativeNumber(result.details.c, 0), 4) + '</td></tr>';
        html += '<tr><td>M (Table 9-3)</td><td class="num">' + formatNumber(parseNonNegativeNumber(result.details.m, 0), 4) + ' lb</td></tr>';
        html += '<tr><td>hc</td><td class="num">' + formatNumber(parseNonNegativeNumber(result.details.hc, 0), 4) + '</td></tr>';
        html += '<tr><td>Mc</td><td class="num">' + formatNumber(parseNonNegativeNumber(result.details.mc, 0), 4) + ' lb</td></tr>';
        html += '<tr><td>Mv</td><td class="num">' + formatNumber(parseNonNegativeNumber(result.details.mv, 0), 4) + ' lb</td></tr>';
        html += '<tr><td>mMAX No Vent</td><td class="num">' + formatNumber(parseNonNegativeNumber(result.details.mmaxNoVent, 0), 4) + ' lb</td></tr>';
        html += '<tr><td>mMAX With Vent</td><td class="num">' + formatNumber(parseNonNegativeNumber(result.details.mmaxWithVent, 0), 4) + ' lb</td></tr>';
        if (result.details.mrel !== undefined) html += '<tr><td>mrel</td><td class="num">' + formatNumber(parseNonNegativeNumber(result.details.mrel, 0), 2) + ' lb</td></tr>';
        html += '<tr><td>Comparison</td><td class="num">' + (result.details.compareLabel || 'ms') + ' = ' + formatNumber(parseNonNegativeNumber(result.details.compareValue, result.charge.ms), 2) + ' lb</td></tr>';
        html += '</tbody></table>';
      }

      var conclusion = 'Additional mitigation/manual review required.';
      if (result.status === 'COMPLIANT') conclusion = 'System charge is within allowable limit.';
      if (result.status === 'NOT COMPLIANT') conclusion = 'System charge exceeds allowable limit.';
      if (result.status === 'INCOMPLETE') conclusion = 'Enter required inputs to evaluate compliance.';
      var footerMessage = result.warning || result.classificationWarning || conclusion;
      html += '<div class="refrig-result-footer">';
      html += '<span class="refrig-conclusion"><strong>' + footerMessage + '</strong></span>';
      html += '<span class="refrig-timestamp">' + (result.timestamp ? ('Calculated at: ' + result.timestamp) : '') + '</span>';
      html += '</div>';
      return html;
    }

    function updateRefrigerantVisibility(state, result) {
      var complianceStatusNode = document.getElementById('refComplianceLiveStatus');
      if (complianceStatusNode) {
        var complianceStatusText = (result && result.status) ? result.status : 'INCOMPLETE';
        var complianceClass = 'incomplete';
        if (complianceStatusText === 'COMPLIANT') complianceClass = 'ok';
        else if (complianceStatusText === 'NOT COMPLIANT') complianceClass = 'fail';
        else if (complianceStatusText === 'MANUAL REVIEW REQUIRED') complianceClass = 'review';
        complianceStatusNode.className = 'refrig-status ' + complianceClass;
        complianceStatusNode.innerHTML = complianceStatusText;
      }

      var calcMode = state.calculationMode || 'SINGLE';
      var singleNode = document.getElementById('refSingleModeContent');
      var multiNode = document.getElementById('refMultiModeContent');
      if (singleNode) singleNode.style.display = calcMode === 'SINGLE' ? 'block' : 'none';
      if (multiNode) multiNode.style.display = calcMode === 'MULTI_ZONE' ? 'block' : 'none';

      if (calcMode === 'MULTI_ZONE') {
        var summaryNode = document.getElementById('refMultiSummary');
        var sum = result && result.summary ? result.summary : { totalRows: 0, compliantRows: 0, nonCompliantRows: 0, manualReviewRows: 0 };
        var body = document.getElementById('refMultiRowsBody');
        var i, tr, rowRes;
        if (summaryNode) {
          summaryNode.innerHTML = 'Rows: ' + parseNonNegativeNumber(sum.totalRows, 0) +
            ' | Compliant: ' + parseNonNegativeNumber(sum.compliantRows, 0) +
            ' | Not Compliant: ' + parseNonNegativeNumber(sum.nonCompliantRows, 0) +
            ' | Manual Review: ' + parseNonNegativeNumber(sum.manualReviewRows, 0);
        }
        if (body && result && result.rows && result.rows.length) {
          for (i = 0; i < body.rows.length; i++) {
            tr = body.rows[i];
            rowRes = result.rows[i];
            if (!tr || !rowRes) continue;
            if (tr.cells[10]) tr.cells[10].innerHTML = formatNumber(parseNonNegativeNumber(rowRes.lineCharge, 0), 4);
            if (tr.cells[11]) tr.cells[11].innerHTML = formatNumber(parseNonNegativeNumber(rowRes.totalCharge, 0), 4);
            if (tr.cells[15]) tr.cells[15].innerHTML = formatNumber(parseNonNegativeNumber(rowRes.c, 0), 4);
            if (tr.cells[16]) tr.cells[16].innerHTML = formatNumber(parseNonNegativeNumber(rowRes.m, 0), 4);
            if (tr.cells[17]) tr.cells[17].innerHTML = formatNumber(parseNonNegativeNumber(rowRes.hc, 0), 4);
            if (tr.cells[18]) tr.cells[18].innerHTML = formatNumber(parseNonNegativeNumber(rowRes.mc, 0), 4);
            if (tr.cells[19]) tr.cells[19].innerHTML = formatNumber(parseNonNegativeNumber(rowRes.mv, 0), 4);
            if (tr.cells[20]) tr.cells[20].innerHTML = formatNumber(parseNonNegativeNumber(rowRes.mmaxNoVent, 0), 4);
            if (tr.cells[21]) tr.cells[21].innerHTML = formatNumber(parseNonNegativeNumber(rowRes.mmaxWithVent, 0), 4);
            if (tr.cells[22]) tr.cells[22].innerHTML = '<span class="refrig-status ' + (rowRes.statusNoVent === 'COMPLIANT' ? 'ok' : (rowRes.statusNoVent === 'NOT COMPLIANT' ? 'fail' : 'review')) + '">' + (rowRes.complianceNoVent || rowRes.statusNoVent) + '</span>';
            if (tr.cells[23]) tr.cells[23].innerHTML = '<span class="refrig-status ' + (rowRes.statusWithVent === 'COMPLIANT' ? 'ok' : (rowRes.statusWithVent === 'NOT COMPLIANT' ? 'fail' : 'review')) + '">' + (rowRes.complianceWithVent || rowRes.statusWithVent) + '</span>';
          }
        }
        if (document.getElementById('refApplicableStandardBadge')) {
          document.getElementById('refApplicableStandardBadge').innerHTML = (result.applicableStandard || 'ASHRAE 15.2');
        }
        if (document.getElementById('refStandardReason')) {
          document.getElementById('refStandardReason').innerHTML = 'Multi-Zone table mode: shared assumptions are applied across all rows.';
        }
        if (document.getElementById('refLiveSummary')) {
          document.getElementById('refLiveSummary').innerHTML =
            '<span class="refrig-live-pill"><strong>Standard:</strong> ' + (result.applicableStandard || '-') + '</span>' +
            '<span class="refrig-live-pill"><strong>Method:</strong> Multi-Zone Table</span>' +
            '<span class="refrig-live-pill"><strong>Refrigerant:</strong> ' + (state.multiZone && state.multiZone.shared ? state.multiZone.shared.refrigerantType : '-') + '</span>' +
            '<span class="refrig-live-pill"><strong>Status:</strong> ' + (result.status || 'INCOMPLETE') + '</span>' +
            '<span class="refrig-live-pill"><strong>Total Rows:</strong> ' + parseNonNegativeNumber(sum.totalRows, 0) + '</span>';
        }
        return;
      }

      var applicableStandard = result.applicableStandard || '';
      var is15 = applicableStandard === 'ASHRAE 15';
      var is152 = applicableStandard === 'ASHRAE 15.2';
      if (document.getElementById('refAshrae15Section')) document.getElementById('refAshrae15Section').style.display = is15 ? '' : 'none';
      if (document.getElementById('refAshrae152Section')) document.getElementById('refAshrae152Section').style.display = is152 ? '' : 'none';

      var isManual = state.ashrae15Inputs.selectionMode === 'MANUAL';
      if (document.getElementById('refManualEqRow')) document.getElementById('refManualEqRow').style.display = is15 && isManual ? 'flex' : 'none';
      var eq = selectAshrae15Equation(state.ashrae15Inputs);
      if (document.getElementById('refEq78Fields')) document.getElementById('refEq78Fields').style.display = (is15 && eq === 'EQ_7_8') ? '' : 'none';
      if (document.getElementById('refEq79Fields')) document.getElementById('refEq79Fields').style.display = (is15 && eq === 'EQ_7_9') ? '' : 'none';
      if (document.getElementById('refEquationBadge')) document.getElementById('refEquationBadge').innerHTML = 'Using ' + (eq === 'EQ_7_9' ? 'Equation 7-9' : 'Equation 7-8');
      if (document.getElementById('refApplicableStandardBadge')) {
        document.getElementById('refApplicableStandardBadge').innerHTML = (applicableStandard || 'Manual review required');
      }
      if (document.getElementById('refStandardReason')) {
        document.getElementById('refStandardReason').innerHTML = result.standardReason || '';
      }

      if (state.classification.refrigerantType === 'R-454B' && document.getElementById('refLfl78')) {
        if (Math.abs(parseFloat(document.getElementById('refLfl78').value || '0')) <= 0.0001) {
          document.getElementById('refLfl78').value = '18.5';
        }
      }
      if (document.getElementById('refManualAdditionalCharge')) {
        document.getElementById('refManualAdditionalCharge').disabled = !state.chargeInputs.useManualAdditionalCharge;
      }
      if (document.getElementById('refManualAdditionalChargeRow')) {
        document.getElementById('refManualAdditionalChargeRow').style.display = state.chargeInputs.useManualAdditionalCharge ? '' : 'none';
      }
      if (document.getElementById('refVentilationCfmRow152')) {
        document.getElementById('refVentilationCfmRow152').style.display = (state.ashrae152Inputs.ventilationUsed === 'YES') ? '' : 'none';
      }
      if (document.getElementById('refCOverrideRow152')) {
        document.getElementById('refCOverrideRow152').style.display = state.ashrae152Inputs.useCOverride ? '' : 'none';
      }
      if (document.getElementById('refCcalc152')) document.getElementById('refCcalc152').value = formatNumber(parseNonNegativeNumber(result.details.c, 0), 4);
      if (document.getElementById('refMcalc152')) document.getElementById('refMcalc152').value = formatNumber(parseNonNegativeNumber(result.details.m, 0), 4);
      if (document.getElementById('refHc152')) document.getElementById('refHc152').value = formatNumber(parseNonNegativeNumber(result.details.hc, 1), 4);
      if (document.getElementById('refMc152')) document.getElementById('refMc152').value = formatNumber(parseNonNegativeNumber(result.details.mc, 0), 4);
      if (document.getElementById('refMv152')) document.getElementById('refMv152').value = formatNumber(parseNonNegativeNumber(result.details.mv, 0), 4);
      if (document.getElementById('refMmaxNoVent152')) document.getElementById('refMmaxNoVent152').value = formatNumber(parseNonNegativeNumber(result.details.mmaxNoVent, 0), 4);
      if (document.getElementById('refMmaxWithVent152')) document.getElementById('refMmaxWithVent152').value = formatNumber(parseNonNegativeNumber(result.details.mmaxWithVent, 0), 4);

      if (document.getElementById('refLiveSummary')) {
        var methodTxt = result.methodLabel || '-';
        var statusTxt = result.status || 'INCOMPLETE';
        var totalChargeTxt = formatNumber(parseNonNegativeNumber((result.charge && result.charge.ms) ? result.charge.ms : 0, 0), 2) + ' pounds';
        document.getElementById('refLiveSummary').innerHTML =
          '<span class="refrig-live-pill"><strong>Standard:</strong> ' + (result.applicableStandard || '-') + '</span>' +
          '<span class="refrig-live-pill"><strong>Method:</strong> ' + methodTxt + '</span>' +
          '<span class="refrig-live-pill"><strong>Refrigerant:</strong> ' + (result.refrigerantType || state.classification.refrigerantType || '-') + '</span>' +
          '<span class="refrig-live-pill"><strong>Status:</strong> ' + statusTxt + '</span>' +
          '<span class="refrig-live-pill"><strong>Total Charge:</strong> ' + totalChargeTxt + '</span>';
      }
    }

    function onRefrigerantInputsChanged() {
      if (document.getElementById('refCalculationMode') && document.getElementById('refCalculationMode').value === 'MULTI_ZONE') {
        if (refrigerantMultiRows.length <= 0) addRefrigerantMultiRow();
        else renderRefrigerantMultiRows();
      }
      calculateRefrigerantCompliance(false);
    }

    function calculateRefrigerantCompliance(commit) {
      var state = getRefrigerantState();
      var result;
      if (state.calculationMode === 'MULTI_ZONE') {
        result = calculateRefrigerantMultiZone(state.multiZone.shared, state.classification);
      } else {
        result = buildRefrigerantComplianceResult(state);
        result.timestamp = new Date().toISOString();
        result.charge.factoryCharge = state.chargeInputs.factoryCharge;
        result.charge.fieldCharge = state.chargeInputs.fieldCharge;
      }
      if (document.getElementById('refChargeSummary') && state.calculationMode !== 'MULTI_ZONE') {
        document.getElementById('refChargeSummary').innerHTML =
          '<span class=\"refrig-charge-item\"><span class=\"refrig-charge-label\">Additional Charge</span><span class=\"refrig-charge-value\">' + formatNumber(parseNonNegativeNumber(result.charge.additionalCharge, 0), 2) + ' lb</span></span>' +
          '<span class=\"refrig-charge-item\"><span class=\"refrig-charge-label\">Total System Charge (ms)</span><span class=\"refrig-charge-value\">' + formatNumber(parseNonNegativeNumber(result.charge.ms, 0), 2) + ' lb</span></span>';
      }
      updateRefrigerantVisibility(state, result);
      var previewText = '';
      if (state.calculationMode === 'MULTI_ZONE') {
        previewText = 'Live table check: ' + parseNonNegativeNumber(result.summary.totalRows, 0) + ' rows evaluated.';
      } else {
        if (result.status === 'INCOMPLETE') previewText = 'Live check: Enter required inputs to evaluate compliance.';
        else if (result.status === 'MANUAL REVIEW REQUIRED') previewText = 'Live check: Manual review path identified.';
        else previewText = 'Live check: ' + result.status + ' (' + (result.methodLabel || '-') + ')';
      }
      if (document.getElementById('refPreview')) document.getElementById('refPreview').innerHTML = previewText;
      if (document.getElementById('refPreviewMulti')) document.getElementById('refPreviewMulti').innerHTML = previewText;

      if (document.getElementById('refrigerantResults')) {
        document.getElementById('refrigerantResults').innerHTML = renderRefrigerantResults(result);
      }
      latestRefrigerantSnapshot = {
        section: 'refrigerant',
        selection: state,
        result: result
      };
      if (commit !== true) return;
    }

    function resetRefrigerantSection() {
      if (document.getElementById('refCalculationMode')) document.getElementById('refCalculationMode').value = 'SINGLE';
      if (document.getElementById('refRefrigerantType')) document.getElementById('refRefrigerantType').value = 'R-454B';
      if (document.getElementById('refSafetyGroup')) document.getElementById('refSafetyGroup').value = 'A2L';
      if (document.getElementById('refSystemType')) document.getElementById('refSystemType').value = 'SPLIT_DX';
      if (document.getElementById('refSpaceType')) document.getElementById('refSpaceType').value = 'RESIDENTIAL';
      if (document.getElementById('refOneSystemOneDwelling')) document.getElementById('refOneSystemOneDwelling').value = 'YES';
      if (document.getElementById('refHighProbability')) document.getElementById('refHighProbability').value = 'YES';
      if (document.getElementById('refDataVersion')) document.getElementById('refDataVersion').value = 'ADDENDA';

      if (document.getElementById('refFactoryCharge')) document.getElementById('refFactoryCharge').value = 0;
      if (document.getElementById('refFieldCharge')) document.getElementById('refFieldCharge').value = 0;
      if (document.getElementById('refInstalledPipeLength')) document.getElementById('refInstalledPipeLength').value = 0;
      if (document.getElementById('refIncludedPipeLength')) document.getElementById('refIncludedPipeLength').value = 0;
      if (document.getElementById('refAdditionalChargeRate')) document.getElementById('refAdditionalChargeRate').value = 0;
      if (document.getElementById('refUseManualAdditionalCharge')) document.getElementById('refUseManualAdditionalCharge').checked = false;
      if (document.getElementById('refManualAdditionalCharge')) document.getElementById('refManualAdditionalCharge').value = 0;

      if (document.getElementById('refEqSelectionMode')) document.getElementById('refEqSelectionMode').value = 'AUTO';
      if (document.getElementById('refConnectedSpaces')) document.getElementById('refConnectedSpaces').value = 'NO';
      if (document.getElementById('refAirCirculation')) document.getElementById('refAirCirculation').value = 'YES';
      if (document.getElementById('refManualEquation')) document.getElementById('refManualEquation').value = 'EQ_7_8';
      if (document.getElementById('refRoomArea78')) document.getElementById('refRoomArea78').value = 0;
      if (document.getElementById('refCeilingHeight78')) document.getElementById('refCeilingHeight78').value = 0;
      if (document.getElementById('refLfl78')) document.getElementById('refLfl78').value = 18.5;
      if (document.getElementById('refCf78')) document.getElementById('refCf78').value = 0.5;
      if (document.getElementById('refFocc78')) document.getElementById('refFocc78').value = 1;
      if (document.getElementById('refMdef79')) document.getElementById('refMdef79').value = 0;
      if (document.getElementById('refFlfl79')) document.getElementById('refFlfl79').value = 1;
      if (document.getElementById('refFocc79')) document.getElementById('refFocc79').value = 1;
      if (document.getElementById('refLowestOpeningHeight79')) document.getElementById('refLowestOpeningHeight79').value = 0;
      if (document.getElementById('refOptionalRoomArea79')) document.getElementById('refOptionalRoomArea79').value = 0;

      if (document.getElementById('refM1_152')) document.getElementById('refM1_152').value = 0;
      if (document.getElementById('refM2_152')) document.getElementById('refM2_152').value = 0;
      if (document.getElementById('refRoomArea152')) document.getElementById('refRoomArea152').value = 0;
      if (document.getElementById('refRoomVolume152')) document.getElementById('refRoomVolume152').value = 0;
      if (document.getElementById('refSpaceHeight152')) document.getElementById('refSpaceHeight152').value = 7.2;
      if (document.getElementById('refVentilationCfm152')) document.getElementById('refVentilationCfm152').value = 0;
      if (document.getElementById('refInterpolate152')) document.getElementById('refInterpolate152').value = 'YES';
      if (document.getElementById('refAllowAreaExtrapolation152')) document.getElementById('refAllowAreaExtrapolation152').value = 'NO';
      if (document.getElementById('refUseCOverride152')) document.getElementById('refUseCOverride152').value = 'NO';
      if (document.getElementById('refCOverride152')) document.getElementById('refCOverride152').value = 0;
      if (document.getElementById('refVentilation152')) document.getElementById('refVentilation152').value = 'NO';
      if (document.getElementById('refSsov152')) document.getElementById('refSsov152').value = 'NO';
      if (document.getElementById('refUseMrel152')) document.getElementById('refUseMrel152').value = 'NO';
      if (document.getElementById('refMrel152')) document.getElementById('refMrel152').value = 0;
      if (document.getElementById('refCcalc152')) document.getElementById('refCcalc152').value = '0.0000';
      if (document.getElementById('refMcalc152')) document.getElementById('refMcalc152').value = '0.0000';
      if (document.getElementById('refHc152')) document.getElementById('refHc152').value = '1.0000';
      if (document.getElementById('refMc152')) document.getElementById('refMc152').value = '0.0000';
      if (document.getElementById('refMv152')) document.getElementById('refMv152').value = '0.0000';
      if (document.getElementById('refMmaxNoVent152')) document.getElementById('refMmaxNoVent152').value = '0.0000';
      if (document.getElementById('refMmaxWithVent152')) document.getElementById('refMmaxWithVent152').value = '0.0000';

      if (document.getElementById('refMzRefrigerantType')) document.getElementById('refMzRefrigerantType').value = 'R-454B';
      if (document.getElementById('refMzSafetyGroup')) document.getElementById('refMzSafetyGroup').value = 'A2L';
      if (document.getElementById('refMzStandardPath')) document.getElementById('refMzStandardPath').value = 'AUTO';
      if (document.getElementById('refMzDataVersion')) document.getElementById('refMzDataVersion').value = 'ADDENDA';
      if (document.getElementById('refMzDefaultPipeSize')) document.getElementById('refMzDefaultPipeSize').value = '3/8 in';
      if (document.getElementById('refMzDefaultChargePerFt')) document.getElementById('refMzDefaultChargePerFt').value = 0;
      if (document.getElementById('refMzDefaultFactoryCharge')) document.getElementById('refMzDefaultFactoryCharge').value = 0;
      if (document.getElementById('refMzVentilationConsidered')) document.getElementById('refMzVentilationConsidered').value = 'NO';
      if (document.getElementById('refMzSsovConsidered')) document.getElementById('refMzSsovConsidered').value = 'NO';
      if (document.getElementById('refMzDefaultUnits')) document.getElementById('refMzDefaultUnits').value = 1;
      if (document.getElementById('refMzDefaultSpaceHeight')) document.getElementById('refMzDefaultSpaceHeight').value = 7.2;
      if (document.getElementById('refMzDefaultVentCfm')) document.getElementById('refMzDefaultVentCfm').value = 0;
      if (document.getElementById('refMzDefaultM1')) document.getElementById('refMzDefaultM1').value = 0;
      if (document.getElementById('refMzDefaultM2')) document.getElementById('refMzDefaultM2').value = 0;
      if (document.getElementById('refMzInterpolate')) document.getElementById('refMzInterpolate').value = 'YES';
      if (document.getElementById('refMzAllowAreaExtrapolation')) document.getElementById('refMzAllowAreaExtrapolation').value = 'NO';
      refrigerantMultiRows = [];
      refrigerantMultiRowIdSeq = 0;
      addRefrigerantMultiRow();
      latestRefrigerantSnapshot = null;
      calculateRefrigerantCompliance(false);
    }

    async function initSolarCatalog() {
      if (!window.pywebview || !window.pywebview.api || !window.pywebview.api.get_solar_catalog) return;
      try {
        var r = await window.pywebview.api.get_solar_catalog();
        var typeNode = document.getElementById('solarBuildingType');
        if (!typeNode) return;
        if (r && r.ok && r.buildingTypes && r.buildingTypes.length > 0) {
          typeNode.options.length = 0;
          var i, opt;
          for (i = 0; i < r.buildingTypes.length; i++) {
            opt = document.createElement('option');
            opt.value = r.buildingTypes[i];
            opt.text = r.buildingTypes[i];
            typeNode.add(opt);
          }
          typeNode.value = r.buildingTypes[0];
        }
      } catch (e) {}
    }

    function getSolarSelectionState() {
      return {
        buildingSf: parseNonNegativeNumber(document.getElementById('solarBuildingSf') ? document.getElementById('solarBuildingSf').value : 0, 0),
        zipCode: document.getElementById('solarZip') ? document.getElementById('solarZip').value : '',
        climateZone: parseNonNegativeNumber(document.getElementById('solarClimateZone') ? document.getElementById('solarClimateZone').value : 0, 0),
        buildingType: document.getElementById('solarBuildingType') ? document.getElementById('solarBuildingType').value : 'Grocery',
        useSaraPath: !!(document.getElementById('solarUseSaraPath') && document.getElementById('solarUseSaraPath').checked),
        sara: parseNonNegativeNumber(document.getElementById('solarSara') ? document.getElementById('solarSara').value : 0, 0),
        dValue: parseNonNegativeNumber(document.getElementById('solarDValue') ? document.getElementById('solarDValue').value : 0, 0),
        solarExempt: !!(document.getElementById('solarExempt') && document.getElementById('solarExempt').checked),
        batteryExempt: !!(document.getElementById('batteryExempt') && document.getElementById('batteryExempt').checked),
        notes: document.getElementById('solarNotes') ? document.getElementById('solarNotes').value : ''
      };
    }

    function setSolarSaraPathUiState(enabled) {
      var saraRow = document.getElementById('solarSaraRow');
      var dRow = document.getElementById('solarDValueRow');
      var saraInput = document.getElementById('solarSara');
      var dInput = document.getElementById('solarDValue');
      if (saraInput) saraInput.disabled = !enabled;
      if (dInput) dInput.disabled = !enabled;
      if (saraRow) saraRow.classList.toggle('solar-path-disabled', !enabled);
      if (dRow) dRow.classList.toggle('solar-path-disabled', !enabled);
    }

    function onSolarSaraToggle() {
      var enabled = !!(document.getElementById('solarUseSaraPath') && document.getElementById('solarUseSaraPath').checked);
      setSolarSaraPathUiState(enabled);
      updateSolarPreview();
    }

    async function lookupSolarClimateFromZip(zip) {
      if (!window.pywebview || !window.pywebview.api || !window.pywebview.api.lookup_solar_climate) return null;
      try {
        var r = await window.pywebview.api.lookup_solar_climate(zip);
        if (r && r.ok && parseNonNegativeNumber(r.climateZone, 0) > 0) return parseNonNegativeNumber(r.climateZone, 0);
      } catch (e) {}
      return null;
    }

    function onSolarClimateZoneInput() {
      var node = document.getElementById('solarClimateZone');
      var val = parseNonNegativeNumber(node ? node.value : 0, 0);
      solarClimateManualOverride = val > 0;
      if (!solarClimateManualOverride) {
        onSolarZipInput();
      }
      updateSolarPreview();
    }

    function onSolarZipInput() {
      var zipNode = document.getElementById('solarZip');
      var zip = zipNode ? String(zipNode.value || '').trim() : '';
      updateSolarPreview();
      if (solarZipLookupTimer) {
        window.clearTimeout(solarZipLookupTimer);
        solarZipLookupTimer = null;
      }
      solarZipLookupTimer = window.setTimeout(async function () {
        var climateNode = document.getElementById('solarClimateZone');
        if (!climateNode) return;
        if (solarClimateManualOverride) return;
        if (!/^\d{5}$/.test(zip)) {
          if (!solarClimateManualOverride) climateNode.value = '';
          return;
        }
        var cz = await lookupSolarClimateFromZip(zip);
        if (cz && !solarClimateManualOverride) climateNode.value = cz;
        updateSolarPreview();
      }, 220);
    }

    function renderSolarResult(r) {
      var compliancePath = r.useSaraPath ? 'SARA Path' : 'CFA Path';
      var batteryStatus = (r.batteryExemptionStatus || 'NOT EXEMPT');
      return '' +
        '<div class="solar-results-section">' +
          '<div class="solar-results-title">Solar Requirement Summary</div>' +
          '<div class="solar-summary-grid">' +
            '<div class="solar-summary-item main"><span class="solar-summary-label">Required PV Size</span><span class="solar-summary-value">' + formatNumber(parseNonNegativeNumber(r.pvKwdc, 0), 2) + ' kW</span></div>' +
            '<div class="solar-summary-item"><span class="solar-summary-label">Minimum SARA Required</span><span class="solar-summary-value">' + formatNumber(parseNonNegativeNumber(r.minimumSara, 0), 2) + ' ft²</span></div>' +
            '<div class="solar-summary-item"><span class="solar-summary-label">Climate Zone</span><span class="solar-summary-value">' + (r.climateZone || '-') + '</span></div>' +
            '<div class="solar-summary-item"><span class="solar-summary-label">Compliance Path Used</span><span class="solar-summary-value">' + compliancePath + '</span></div>' +
          '</div>' +
        '</div>' +
        '<div class="solar-results-section">' +
          '<div class="solar-results-title">Battery Requirement</div>' +
          '<div class="solar-battery-grid">' +
            '<div class="solar-summary-item"><span class="solar-summary-label">Required Battery Capacity</span><span class="solar-summary-value">' + formatNumber(parseNonNegativeNumber(r.batteryKwh, 0), 2) + ' kWh</span></div>' +
            '<div class="solar-summary-item"><span class="solar-summary-label">Battery Exemption Status</span><span class="solar-summary-value">' + batteryStatus + '</span></div>' +
          '</div>' +
        '</div>' +
        '<div class="solar-results-section">' +
          '<div class="solar-results-title">Notes / Code Reference</div>' +
          '<div class="solar-note-box">' + (r.message || 'Workbook factors and selected compliance path were used for this worksheet result.') + '</div>' +
        '</div>';
    }

    function updateSolarPreview() {
      var node = document.getElementById('solarPreview');
      if (!node) return;
      var sel = getSolarSelectionState();
      if (sel.buildingSf <= 0) {
        node.innerHTML = 'Enter Building SF > 0 to preview.';
        scheduleSolarAutoCalculation();
        return;
      }
      var dValue = sel.dValue > 0 ? sel.dValue : 0.95;
      var pvKw = (sel.buildingSf * 0.44) / 1000.0;
      var batteryKwh = (pvKw * 0.93) / Math.sqrt(dValue);
      var batteryKw = pvKw * 0.23;
      node.innerHTML =
        'Preview: PV ' + formatNumber(pvKw, 2) + ' kWdc | Battery ' +
        formatNumber(batteryKwh, 2) + ' kWh | ' + formatNumber(batteryKw, 2) + ' kW' +
        ' | Path ' + (sel.useSaraPath ? 'SARA' : 'CFA');
      scheduleSolarAutoCalculation();
    }

    function scheduleSolarAutoCalculation() {
      if (solarAutoCalcTimer) {
        window.clearTimeout(solarAutoCalcTimer);
        solarAutoCalcTimer = null;
      }
      solarAutoCalcTimer = window.setTimeout(async function () {
        if (activeModule !== 'solar') return;
        var sel = getSolarSelectionState();
        var resultNode = document.getElementById('solarResults');
        if (sel.buildingSf <= 0) {
          latestSolarSnapshot = null;
          if (resultNode) resultNode.innerHTML = '<div class="result-placeholder">Enter worksheet inputs to see live Solar sizing results.</div>';
          return;
        }
        var zip = String(sel.zipCode || '').trim();
        var hasClimateInput = parseNonNegativeNumber(sel.climateZone, 0) > 0 || /^\d{5}$/.test(zip);
        if (!hasClimateInput) {
          latestSolarSnapshot = null;
          if (resultNode) resultNode.innerHTML = '<div class="result-placeholder">Enter ZIP or Climate Zone to evaluate solar compliance outputs.</div>';
          return;
        }
        await calculateSolar(true);
      }, 280);
    }

    async function calculateSolar(isAutoMode) {
      var autoMode = !!isAutoMode;
      var sel = getSolarSelectionState();
      var resultNode = document.getElementById('solarResults');
      if (sel.buildingSf <= 0) {
        latestSolarSnapshot = null;
        if (autoMode) {
          if (resultNode) resultNode.innerHTML = '<div class="result-placeholder">Enter Building Total SF (CFA) to evaluate solar sizing.</div>';
        } else {
          setWarningResult(resultNode, 'Enter Building Total SF (CFA) greater than zero.');
        }
        return;
      }
      if (!window.pywebview || !window.pywebview.api || !window.pywebview.api.calculate_solar) {
        latestSolarSnapshot = null;
        if (autoMode) {
          if (resultNode) resultNode.innerHTML = '<div class="result-placeholder">Solar service is not ready yet.</div>';
        } else {
          setWarningResult(resultNode, 'Solar service is not ready yet. Please try again.');
        }
        return;
      }
      try {
        var r = await window.pywebview.api.calculate_solar(sel);
        if (!r || !r.ok) {
          latestSolarSnapshot = null;
          if (autoMode) {
            if (resultNode) resultNode.innerHTML = '<div class="result-placeholder">' + (r && r.reason ? r.reason : 'Complete required inputs to calculate.') + '</div>';
          } else {
            setWarningResult(resultNode, (r && r.reason ? r.reason : 'Solar calculation failed.'));
          }
          return;
        }
        if (document.getElementById('solarClimateZone') && !solarClimateManualOverride) document.getElementById('solarClimateZone').value = r.climateZone;
        latestSolarSnapshot = {
          section: 'solar',
          selection: getSolarSelectionState(),
          result: r
        };
        if (resultNode) {
          resultNode.innerHTML = renderSolarResult(r);
        }
      } catch (e) {
        latestSolarSnapshot = null;
        if (autoMode) {
          if (resultNode) resultNode.innerHTML = '<div class="result-placeholder">Enter required inputs to evaluate live solar results.</div>';
        } else {
          setWarningResult(resultNode, 'Solar calculation failed. ' + ((e && e.message) ? e.message : 'Unknown error'));
        }
      }
    }

    // ---- Section-specific zone architecture overrides (Sanitary / WSFU / Duct) ----
    function makeSanitaryZoneId() {
      sanitaryZoneIdSeq += 1;
      return 'sd_zone_' + sanitaryZoneIdSeq;
    }

    function makeSanitaryFixtureId() {
      ventFixtureIdSeq += 1;
      return 'sd_fixture_' + ventFixtureIdSeq;
    }

    function makeWsfuZoneId() {
      wsfuZoneIdSeq += 1;
      return 'wsfu_zone_' + wsfuZoneIdSeq;
    }

    function makeDuctZoneId() {
      ductZoneIdSeq += 1;
      return 'duct_zone_' + ductZoneIdSeq;
    }

    function makeDuctSegmentId() {
      ductSegmentIdSeq += 1;
      return 'duct_seg_' + ductSegmentIdSeq;
    }

    function createDefaultSanitaryFixtureRow() {
      return {
        id: makeSanitaryFixtureId(),
        fixtureKey: (ventFixtureCatalog.length > 0 ? ventFixtureCatalog[0].key : ''),
        quantity: 0,
        unitDfu: 0,
        rowTotalDfu: 0,
        fixtureLabel: ''
      };
    }

    function createDefaultSanitaryZone(indexHint) {
      var index = parseInt(indexHint, 10);
      if (isNaN(index) || index < 1) index = sanitaryZones.length + 1;
      return {
        id: makeSanitaryZoneId(),
        name: 'Zone ' + index,
        fixtures: [createDefaultSanitaryFixtureRow()],
        results: {}
      };
    }

    function createDefaultWsfuFixtureRow() {
      return {
        id: makeWsfuFixtureId(),
        fixtureKey: (wsfuFixtureCatalog.length > 0 ? wsfuFixtureCatalog[0].key : ''),
        quantity: 0,
        hotCold: false,
        unitWsfu: 0,
        rowTotalWsfu: 0,
        fixtureLabel: ''
      };
    }

    function createDefaultWsfuZone(indexHint) {
      var index = parseInt(indexHint, 10);
      if (isNaN(index) || index < 1) index = wsfuZones.length + 1;
      return {
        id: makeWsfuZoneId(),
        name: 'Zone ' + index,
        fixtures: [createDefaultWsfuFixtureRow()],
        results: {}
      };
    }

    function createDefaultDuctSegment() {
      return {
        id: makeDuctSegmentId(),
        name: 'Segment ' + (ductSegmentIdSeq || 1),
        mode: DUCT_CALC.modes.SIZE_FROM_CFM_FRICTION,
        shape: DUCT_CALC.shapes.ROUND,
        cfm: 0,
        frictionTarget: 0,
        velocityTarget: 0,
        diameter: 0,
        width: 0,
        height: 0,
        ratioLimit: 3
      };
    }

    function createDefaultDuctZone(indexHint) {
      var index = parseInt(indexHint, 10);
      if (isNaN(index) || index < 1) index = ductZones.length + 1;
      return {
        id: makeDuctZoneId(),
        name: 'Zone ' + index,
        segments: [createDefaultDuctSegment()],
        results: {}
      };
    }

    function ensureSanitaryZonesInitialized() {
      if (!sanitaryZones || sanitaryZones.length <= 0) sanitaryZones = [createDefaultSanitaryZone(1)];
    }

    function ensureWsfuZonesInitialized() {
      if (!wsfuZones || wsfuZones.length <= 0) wsfuZones = [createDefaultWsfuZone(1)];
    }

    function ensureDuctZonesInitialized() {
      if (!ductZones || ductZones.length <= 0) ductZones = [createDefaultDuctZone(1)];
    }

    function addSanitaryZone(seedZone) {
      var zone = createDefaultSanitaryZone(sanitaryZones.length + 1);
      var rows, i, row;
      if (seedZone && seedZone.id) zone.id = String(seedZone.id);
      if (seedZone && seedZone.name) zone.name = String(seedZone.name);
      zone.fixtures = [];
      rows = (seedZone && (seedZone.fixtures || seedZone.fixtureRows) && (seedZone.fixtures || seedZone.fixtureRows).length) ? (seedZone.fixtures || seedZone.fixtureRows) : [createDefaultSanitaryFixtureRow()];
      for (i = 0; i < rows.length; i++) {
        row = rows[i] || {};
        zone.fixtures.push({
          id: row.id || makeSanitaryFixtureId(),
          fixtureKey: row.fixtureKey ? normalizeCpcFixtureKey(row.fixtureKey) : (ventFixtureCatalog.length > 0 ? ventFixtureCatalog[0].key : ''),
          quantity: parseNonNegativeNumber(row.quantity, 0),
          unitDfu: parseNonNegativeNumber(row.unitDfu, 0),
          rowTotalDfu: parseNonNegativeNumber(row.rowTotalDfu, 0),
          fixtureLabel: row.fixtureLabel || ''
        });
      }
      if (zone.fixtures.length <= 0) zone.fixtures.push(createDefaultSanitaryFixtureRow());
      sanitaryZones.push(zone);
      renderSanitaryZones();
      updateVentTotalsAndUI();
    }

    function removeSanitaryZone(zoneId) {
      var i;
      if (sanitaryZones.length <= 1) {
        sanitaryZones = [createDefaultSanitaryZone(1)];
      } else {
        for (i = sanitaryZones.length - 1; i >= 0; i--) {
          if (sanitaryZones[i].id === zoneId) sanitaryZones.splice(i, 1);
        }
      }
      renderSanitaryZones();
      updateVentTotalsAndUI();
    }

    function addSanitaryFixtureRow(zoneId, seed) {
      var i, zone = null;
      for (i = 0; i < sanitaryZones.length; i++) if (sanitaryZones[i].id === zoneId) zone = sanitaryZones[i];
      if (!zone) return;
      zone.fixtures.push({
        id: (seed && seed.id) ? seed.id : makeSanitaryFixtureId(),
        fixtureKey: (seed && seed.fixtureKey) ? seed.fixtureKey : (ventFixtureCatalog.length > 0 ? ventFixtureCatalog[0].key : ''),
        quantity: parseNonNegativeNumber(seed && seed.quantity, 0),
        unitDfu: parseNonNegativeNumber(seed && seed.unitDfu, 0),
        rowTotalDfu: parseNonNegativeNumber(seed && seed.rowTotalDfu, 0),
        fixtureLabel: (seed && seed.fixtureLabel) ? seed.fixtureLabel : ''
      });
      renderSanitaryZones();
      updateVentTotalsAndUI();
    }

    function removeSanitaryFixtureRow(zoneId, rowId) {
      var i, j, zone = null;
      for (i = 0; i < sanitaryZones.length; i++) if (sanitaryZones[i].id === zoneId) zone = sanitaryZones[i];
      if (!zone) return;
      if (zone.fixtures.length <= 1) return;
      for (j = zone.fixtures.length - 1; j >= 0; j--) if (zone.fixtures[j].id === rowId) zone.fixtures.splice(j, 1);
      renderSanitaryZones();
      updateVentTotalsAndUI();
    }

    function onSanitaryZoneFieldChange(zoneId, value) {
      var i;
      for (i = 0; i < sanitaryZones.length; i++) if (sanitaryZones[i].id === zoneId) sanitaryZones[i].name = String(value || '');
    }

    function onSanitaryFixtureFieldChange(zoneId, rowId, field, value) {
      var i, j, zone, row;
      for (i = 0; i < sanitaryZones.length; i++) {
        if (sanitaryZones[i].id !== zoneId) continue;
        zone = sanitaryZones[i];
        for (j = 0; j < zone.fixtures.length; j++) {
          row = zone.fixtures[j];
          if (row.id !== rowId) continue;
          if (field === 'fixtureKey') row.fixtureKey = normalizeCpcFixtureKey(value || row.fixtureKey);
          if (field === 'quantity') row.quantity = parseNonNegativeNumber(value, 0);
          break;
        }
      }
      updateVentTotalsAndUI();
    }

    function renderSanitaryZones() {
      var node = document.getElementById('ventFixtureRows');
      var html = '';
      var i, j, zone, row, usageType, item, options, unit;
      if (!node) return;
      ensureSanitaryZonesInitialized();
      usageType = getVentUsageType();
      for (i = 0; i < sanitaryZones.length; i++) {
        zone = sanitaryZones[i];
        var zoneFixtures = (zone.fixtures || []).length;
        var zoneDfu = parseNonNegativeNumber((zone.results && zone.results.zoneTotalDfu) ? zone.results.zoneTotalDfu : 0, 0);
        html += '<div class="cond-zone-card">';
        html += '<div class="cond-zone-header"><div class="cond-zone-title">Zone ' + (i + 1) + '</div><div class="cond-zone-name-wrap"><span class="cond-zone-name-label">Zone Name</span><input type="text" value="' + (zone.name || ('Zone ' + (i + 1))) + '" oninput="onSanitaryZoneFieldChange(\'' + zone.id + '\', this.value);" /></div><div class="zone-subtitle">' + zoneFixtures + ' fixtures | Zone DFU: ' + formatNumber(zoneDfu, 2) + '</div></div>';
        html += '<div class="cond-zone-body"><table class="vent-fixture-table"><colgroup><col class="col-fixture" /><col class="col-qty" /><col class="col-unit" /><col class="col-row" /><col class="col-action" /></colgroup><thead><tr><th>Fixture</th><th class="num">Qty</th><th class="num">Unit DFU</th><th class="num">Row DFU</th><th></th></tr></thead><tbody>';
        for (j = 0; j < zone.fixtures.length; j++) {
          row = zone.fixtures[j];
          options = '';
          for (var k = 0; k < ventFixtureCatalog.length; k++) {
            item = ventFixtureCatalog[k];
            unit = ventFixtureDfuByUsage(item, usageType);
            if (unit === null) continue;
            options += '<option value="' + item.key + '"' + (item.key === normalizeCpcFixtureKey(row.fixtureKey) ? ' selected' : '') + '>' + item.label + '</option>';
          }
          if (!options) options = '<option value="">No fixtures available</option>';
          html += '<tr><td><select style="width:100%;" onchange="onSanitaryFixtureFieldChange(\'' + zone.id + '\', \'' + row.id + '\', \'fixtureKey\', this.value);">' + options + '</select></td><td class="num"><input type="number" step="1" min="0" class="vent-qty-input" value="' + parseNonNegativeNumber(row.quantity, 0) + '" oninput="onSanitaryFixtureFieldChange(\'' + zone.id + '\', \'' + row.id + '\', \'quantity\', this.value);" /></td><td class="num"><span class="vent-dfu-badge">' + formatNumber(parseNonNegativeNumber(row.unitDfu, 0), 2) + '</span></td><td class="num"><span class="vent-dfu-badge">' + formatNumber(parseNonNegativeNumber(row.rowTotalDfu, 0), 2) + '</span></td><td class="num"><input type="button" class="btn btn-small vent-remove-btn" value="Remove" onclick="removeSanitaryFixtureRow(\'' + zone.id + '\', \'' + row.id + '\');" /></td></tr>';
        }
        html += '</tbody></table></div>';
        html += '<div class="cond-zone-actions"><div class="cond-zone-actions-left"><input type="button" class="btn cond-secondary-btn" value="+ ADD FIXTURE" onclick="addSanitaryFixtureRow(\'' + zone.id + '\');" /></div><div class="cond-zone-actions-right"><input type="button" class="btn btn-small zone-danger-btn" value="Remove Zone" onclick="removeSanitaryZone(\'' + zone.id + '\');" /></div></div>';
        html += '</div>';
      }
      node.innerHTML = html;
    }

    function getSanitaryZonesState() {
      var zones = [];
      var i, j, zone, row;
      ensureSanitaryZonesInitialized();
      for (i = 0; i < sanitaryZones.length; i++) {
        zone = sanitaryZones[i];
        var fixtureRows = [];
        for (j = 0; j < zone.fixtures.length; j++) {
          row = zone.fixtures[j];
          fixtureRows.push({
            id: row.id,
            fixtureKey: normalizeCpcFixtureKey(row.fixtureKey),
            fixtureLabel: row.fixtureLabel || '',
            quantity: parseNonNegativeNumber(row.quantity, 0),
            unitDfu: parseNonNegativeNumber(row.unitDfu, 0),
            rowTotalDfu: parseNonNegativeNumber(row.rowTotalDfu, 0)
          });
        }
        zones.push({ id: zone.id, name: zone.name || ('Zone ' + (i + 1)), fixtures: fixtureRows, results: zone.results || {} });
      }
      return zones;
    }

    function computeVentRowsAndTotals() {
      var usageType = getVentUsageType();
      var autoTotal = 0;
      var i, j, zone, row, item, unit;
      ensureSanitaryZonesInitialized();
      for (i = 0; i < sanitaryZones.length; i++) {
        zone = sanitaryZones[i];
        var zoneTotal = 0;
        for (j = 0; j < zone.fixtures.length; j++) {
          row = zone.fixtures[j];
          item = getVentFixtureByKey(row.fixtureKey);
          unit = ventFixtureDfuByUsage(item, usageType);
          if (unit === null) unit = 0;
          row.unitDfu = unit;
          row.rowTotalDfu = roundNumber(parseNonNegativeNumber(row.quantity, 0) * parseNonNegativeNumber(unit, 0), 3);
          row.fixtureLabel = item ? item.label : '';
          zoneTotal += row.rowTotalDfu;
        }
        zone.results = zone.results || {};
        zone.results.zoneTotalDfu = roundNumber(zoneTotal, 3);
        autoTotal += zoneTotal;
      }
      autoTotal = roundNumber(autoTotal, 3);
      if (document.getElementById('ventAutoTotalDfu')) document.getElementById('ventAutoTotalDfu').value = autoTotal;
      return autoTotal;
    }

    function getVentFixtureRowsState() {
      var zones = getSanitaryZonesState();
      var rows = [];
      for (var i = 0; i < zones.length; i++) rows = rows.concat(zones[i].fixtures || []);
      return rows;
    }

    function updateVentTotalsAndUI() {
      var autoTotal = computeVentRowsAndTotals();
      var useManual = !!(document.getElementById('ventUseManualTotal') && document.getElementById('ventUseManualTotal').checked);
      var manualNode = document.getElementById('ventManualTotalDfu');
      var effectiveNode = document.getElementById('ventTotalDfu');
      var effectiveTotal = useManual ? parseNonNegativeNumber(manualNode ? manualNode.value : 0, 0) : autoTotal;
      if (manualNode) manualNode.disabled = !useManual;
      if (effectiveNode) effectiveNode.value = roundNumber(effectiveTotal, 3);
      renderSanitaryZones();
      updateVentPreview();
    }

    function addVentFixtureRow(seed) {
      ensureSanitaryZonesInitialized();
      addSanitaryFixtureRow(sanitaryZones[0].id, seed || null);
    }

    function removeVentFixtureRow(id) {
      var i, j;
      ensureSanitaryZonesInitialized();
      for (i = 0; i < sanitaryZones.length; i++) {
        for (j = 0; j < sanitaryZones[i].fixtures.length; j++) {
          if (sanitaryZones[i].fixtures[j].id === id) return removeSanitaryFixtureRow(sanitaryZones[i].id, id);
        }
      }
    }

    function renderVentResultHtml(snapshot) {
      var zones = (snapshot && snapshot.zones && snapshot.zones.length) ? snapshot.zones : [];
      var i, j, zone, row, rowsHtml, html = '';
      html += '<div class="vent-result-summary"><div class="vent-result-metric"><span class="vent-result-metric-label">Effective DFU</span><span class="vent-result-metric-value">' + formatNumber(parseNonNegativeNumber(snapshot.dfu, 0), 2) + '</span></div><div class="vent-result-metric vent-result-metric-main"><span class="vent-result-metric-label">Recommended Pipe Size</span><span class="vent-result-metric-value">' + (snapshot.recommendedSize || 'Out of range') + '</span></div><div class="vent-result-metric"><span class="vent-result-metric-label">Zones</span><span class="vent-result-metric-value">' + zones.length + '</span></div><div class="vent-result-metric"><span class="vent-result-metric-label">Orientation</span><span class="vent-result-metric-value">' + ((snapshot.drainageOrientation === 'VERTICAL') ? 'Vertical' : 'Horizontal') + '</span></div></div>';
      if (snapshot.warning) html += '<div class="warning-line">' + snapshot.warning + '</div>';
      for (i = 0; i < zones.length; i++) {
        zone = zones[i];
        rowsHtml = '';
        for (j = 0; j < (zone.rows || []).length; j++) {
          row = zone.rows[j];
          rowsHtml += '<tr><td>' + (row.fixtureLabel || row.fixtureKey || '') + '</td><td class="num">' + formatNumber(parseNonNegativeNumber(row.quantity, 0), 0) + '</td><td class="num">' + formatNumber(parseNonNegativeNumber(row.unitDfu, 0), 2) + '</td><td class="num">' + formatNumber(parseNonNegativeNumber(row.rowTotalDfu, 0), 2) + '</td></tr>';
        }
        html += '<div class="cond-zone-card"><div class="cond-zone-header"><div class="cond-zone-title">' + (zone.name || ('Zone ' + (i + 1))) + '</div><div class="module-note">Zone DFU: ' + formatNumber(parseNonNegativeNumber(zone.zoneDfu, 0), 2) + ' | Size: ' + (zone.recommendedSize || 'Out of range') + '</div></div><div class="cond-zone-body"><table class="vent-result-table"><thead><tr><th>Fixture</th><th class="num">Qty</th><th class="num">Unit DFU</th><th class="num">Row DFU</th></tr></thead><tbody>' + (rowsHtml || '<tr><td colspan="4" class="subtle">No fixtures.</td></tr>') + '</tbody></table></div></div>';
      }
      return html;
    }

    function calculateVent() {
      var codeBasis = document.getElementById('ventCodeBasis').value;
      var usageType = getVentUsageType();
      var drainageOrientation = getVentDrainageOrientation();
      var dfu = parseNonNegativeNumber(document.getElementById('ventTotalDfu') ? document.getElementById('ventTotalDfu').value : 0, 0);
      var useManual = !!(document.getElementById('ventUseManualTotal') && document.getElementById('ventUseManualTotal').checked);
      var autoTotal = parseNonNegativeNumber(document.getElementById('ventAutoTotalDfu') ? document.getElementById('ventAutoTotalDfu').value : 0, 0);
      var manualTotal = parseNonNegativeNumber(document.getElementById('ventManualTotalDfu') ? document.getElementById('ventManualTotalDfu').value : 0, 0);
      var resultNode = document.getElementById('ventResults');
      var sizeMatch = findSanitary7032ByDfu(dfu, drainageOrientation);
      var zones = getSanitaryZonesState();
      var zoneResults = [];
      for (var i = 0; i < zones.length; i++) {
        var z = zones[i];
        var zoneDfu = 0;
        for (var j = 0; j < z.fixtures.length; j++) zoneDfu += parseNonNegativeNumber(z.fixtures[j].rowTotalDfu, 0);
        var zMatch = findSanitary7032ByDfu(zoneDfu, drainageOrientation);
        zoneResults.push({ id: z.id, name: z.name, zoneDfu: roundNumber(zoneDfu, 3), recommendedSize: zMatch ? zMatch.size : '', warning: zMatch ? '' : ('Zone ' + (i + 1) + ' exceeds Table 703.2 range.'), rows: z.fixtures });
      }
      if (isNaN(dfu) || dfu <= 0) {
        latestVentSnapshot = null;
        setWarningResult(resultNode, 'Enter an effective DFU value greater than zero.');
        return;
      }
      latestVentSnapshot = {
        section: 'vent',
        codeBasis: codeBasis,
        usageType: usageType,
        drainageOrientation: drainageOrientation,
        dfu: roundNumber(dfu, 3),
        recommendedSize: sizeMatch ? sizeMatch.size : '',
        sizingTable: 'TABLE 703.2',
        tableMaxUnits: sizeMatch ? parseNonNegativeNumber(sizeMatch.maxUnits, 0) : 0,
        maxLengthReference: sizeMatch ? String(sizeMatch.maxLength || 'N/A') : 'N/A',
        warning: sizeMatch ? '' : 'Effective DFU exceeds Table 703.2 maximum for selected orientation.',
        autoTotalDfu: roundNumber(autoTotal, 3),
        manualTotalDfu: roundNumber(manualTotal, 3),
        useManualTotal: useManual,
        rows: getVentFixtureRowsState(),
        zones: zoneResults
      };
      if (resultNode) resultNode.innerHTML = renderVentResultHtml(latestVentSnapshot);
    }

    function resetVentSection() {
      latestVentSnapshot = null;
      if (document.getElementById('ventCodeBasis')) document.getElementById('ventCodeBasis').value = 'IPC';
      if (document.getElementById('ventUsageType')) document.getElementById('ventUsageType').value = 'PRIVATE';
      if (document.getElementById('ventDrainageOrientation')) document.getElementById('ventDrainageOrientation').value = 'HORIZONTAL';
      if (document.getElementById('ventUseManualTotal')) document.getElementById('ventUseManualTotal').checked = false;
      if (document.getElementById('ventManualTotalDfu')) document.getElementById('ventManualTotalDfu').value = 0;
      if (document.getElementById('ventAutoTotalDfu')) document.getElementById('ventAutoTotalDfu').value = 0;
      if (document.getElementById('ventTotalDfu')) document.getElementById('ventTotalDfu').value = 0;
      sanitaryZones = [createDefaultSanitaryZone(1)];
      renderSanitaryZones();
      updateVentTotalsAndUI();
      if (document.getElementById('ventResults')) document.getElementById('ventResults').innerHTML = 'Enter fixture rows and press Calculate.';
    }

    function addWsfuZone(seedZone) {
      var zone = createDefaultWsfuZone(wsfuZones.length + 1);
      var rows, i, row;
      if (seedZone && seedZone.id) zone.id = String(seedZone.id);
      if (seedZone && seedZone.name) zone.name = String(seedZone.name);
      zone.fixtures = [];
      rows = (seedZone && (seedZone.fixtures || seedZone.fixtureRows) && (seedZone.fixtures || seedZone.fixtureRows).length) ? (seedZone.fixtures || seedZone.fixtureRows) : [createDefaultWsfuFixtureRow()];
      for (i = 0; i < rows.length; i++) {
        row = rows[i] || {};
        zone.fixtures.push({
          id: row.id || makeWsfuFixtureId(),
          fixtureKey: row.fixtureKey ? normalizeCpcFixtureKey(row.fixtureKey) : (wsfuFixtureCatalog.length > 0 ? wsfuFixtureCatalog[0].key : ''),
          quantity: parseNonNegativeNumber(row.quantity, 0),
          hotCold: !!row.hotCold,
          unitWsfu: parseNonNegativeNumber(row.unitWsfu, 0),
          rowTotalWsfu: parseNonNegativeNumber(row.rowTotalWsfu, 0),
          fixtureLabel: row.fixtureLabel || ''
        });
      }
      if (zone.fixtures.length <= 0) zone.fixtures.push(createDefaultWsfuFixtureRow());
      wsfuZones.push(zone);
      renderWsfuZones();
      updateWsfuTotalsAndUI();
    }

    function removeWsfuZone(zoneId) {
      if (wsfuZones.length <= 1) wsfuZones = [createDefaultWsfuZone(1)];
      else wsfuZones = wsfuZones.filter(function (z) { return z.id !== zoneId; });
      renderWsfuZones();
      updateWsfuTotalsAndUI();
    }

    function addWsfuFixtureRow(zoneId, seed) {
      var i, zone = null;
      ensureWsfuZonesInitialized();
      for (i = 0; i < wsfuZones.length; i++) if (wsfuZones[i].id === zoneId) zone = wsfuZones[i];
      if (!zone) zone = wsfuZones[0];
      zone.fixtures.push({
        id: (seed && seed.id) ? seed.id : makeWsfuFixtureId(),
        fixtureKey: (seed && seed.fixtureKey) ? seed.fixtureKey : (wsfuFixtureCatalog.length > 0 ? wsfuFixtureCatalog[0].key : ''),
        quantity: parseNonNegativeNumber(seed && seed.quantity, 0),
        hotCold: !!(seed && seed.hotCold),
        unitWsfu: parseNonNegativeNumber(seed && seed.unitWsfu, 0),
        rowTotalWsfu: parseNonNegativeNumber(seed && seed.rowTotalWsfu, 0),
        fixtureLabel: (seed && seed.fixtureLabel) ? seed.fixtureLabel : ''
      });
      renderWsfuZones();
      updateWsfuTotalsAndUI();
    }

    function removeWsfuFixtureRow(zoneId, rowId) {
      var i, j, zone = null;
      for (i = 0; i < wsfuZones.length; i++) if (wsfuZones[i].id === zoneId) zone = wsfuZones[i];
      if (!zone) return;
      if (zone.fixtures.length <= 1) return;
      for (j = zone.fixtures.length - 1; j >= 0; j--) if (zone.fixtures[j].id === rowId) zone.fixtures.splice(j, 1);
      renderWsfuZones();
      updateWsfuTotalsAndUI();
    }

    function onWsfuZoneFieldChange(zoneId, value) {
      for (var i = 0; i < wsfuZones.length; i++) if (wsfuZones[i].id === zoneId) wsfuZones[i].name = String(value || '');
    }

    function onWsfuFixtureFieldChange(zoneId, rowId, field, value) {
      var i, j, zone, row;
      for (i = 0; i < wsfuZones.length; i++) {
        if (wsfuZones[i].id !== zoneId) continue;
        zone = wsfuZones[i];
        for (j = 0; j < zone.fixtures.length; j++) {
          row = zone.fixtures[j];
          if (row.id !== rowId) continue;
          if (field === 'fixtureKey') row.fixtureKey = normalizeCpcFixtureKey(value);
          if (field === 'quantity') row.quantity = parseNonNegativeNumber(value, 0);
          if (field === 'hotCold') row.hotCold = !!value;
        }
      }
      updateWsfuTotalsAndUI();
    }

    function renderWsfuZones() {
      var node = document.getElementById('wsfuFixtureRows');
      var html = '';
      var usageType = document.getElementById('wsfuUsageType') ? document.getElementById('wsfuUsageType').value : 'PRIVATE';
      ensureWsfuZonesInitialized();
      if (!node) return;
      for (var i = 0; i < wsfuZones.length; i++) {
        var zone = wsfuZones[i];
        var zoneCount = (zone.fixtures || []).length;
        var zoneTotal = parseNonNegativeNumber((zone.results && zone.results.zoneTotalFu) ? zone.results.zoneTotalFu : 0, 0);
        html += '<div class="cond-zone-card"><div class="cond-zone-header"><div class="cond-zone-title">Zone ' + (i + 1) + '</div><div class="cond-zone-name-wrap"><span class="cond-zone-name-label">Zone Name</span><input type="text" value="' + (zone.name || ('Zone ' + (i + 1))) + '" oninput="onWsfuZoneFieldChange(\'' + zone.id + '\', this.value);" /></div><div class="zone-subtitle">' + zoneCount + ' fixtures | Zone WSFU: ' + formatNumber(zoneTotal, 2) + '</div></div><div class="cond-zone-body">';
        html += '<table class="wsfu-fixture-table"><thead><tr><th>Fixture</th><th class="num">Qty</th><th class="center">Hot+Cold</th><th class="num">Unit WSFU</th><th class="num">Row Total</th><th class="num">Action</th></tr></thead><tbody>';
        for (var j = 0; j < zone.fixtures.length; j++) {
          var row = zone.fixtures[j];
          var options = '';
          for (var k = 0; k < wsfuFixtureCatalog.length; k++) {
            var item = wsfuFixtureCatalog[k];
            var unit = rawWsfuFixtureUnitByUsage(item, usageType);
            options += '<option value="' + item.key + '"' + (item.key === normalizeCpcFixtureKey(row.fixtureKey) ? ' selected' : '') + '>' + item.label + (unit === null ? ' (N/A)' : '') + '</option>';
          }
          html += '<tr><td><select class="wsfu-fixture-select" onchange="onWsfuFixtureFieldChange(\'' + zone.id + '\', \'' + row.id + '\', \'fixtureKey\', this.value);">' + options + '</select></td><td class="num"><input type="number" min="0" step="1" value="' + parseNonNegativeNumber(row.quantity, 0) + '" class="wsfu-qty-input" oninput="onWsfuFixtureFieldChange(\'' + zone.id + '\', \'' + row.id + '\', \'quantity\', this.value);" /></td><td class="center"><input type="checkbox" class="wsfu-hotcold-checkbox" ' + (row.hotCold ? 'checked' : '') + ' onchange="onWsfuFixtureFieldChange(\'' + zone.id + '\', \'' + row.id + '\', \'hotCold\', this.checked);" /></td><td class="num wsfu-cell-metric">' + formatNumber(parseNonNegativeNumber(row.unitWsfu, 0), 3) + '</td><td class="num wsfu-cell-metric">' + formatNumber(parseNonNegativeNumber(row.rowTotalWsfu, 0), 3) + '</td><td class="num"><input type="button" class="btn btn-remove-row wsfu-remove-btn" value="Remove" onclick="removeWsfuFixtureRow(\'' + zone.id + '\', \'' + row.id + '\');" /></td></tr>';
        }
        html += '</tbody></table></div>';
        html += '<div class="cond-zone-actions"><div class="cond-zone-actions-left"><input type="button" class="btn wsfu-add-fixture-btn" value="+ ADD FIXTURE" onclick="addWsfuFixtureRow(\'' + zone.id + '\');" /></div><div class="cond-zone-actions-right"><input type="button" class="btn btn-small zone-danger-btn" value="Remove Zone" onclick="removeWsfuZone(\'' + zone.id + '\');" /></div></div></div>';
      }
      node.innerHTML = html;
    }

    function getWsfuZonesState() {
      var zones = [];
      ensureWsfuZonesInitialized();
      for (var i = 0; i < wsfuZones.length; i++) {
        zones.push({ id: wsfuZones[i].id, name: wsfuZones[i].name, fixtures: (wsfuZones[i].fixtures || []).map(function (r) { return { id: r.id, fixtureKey: normalizeCpcFixtureKey(r.fixtureKey), fixtureLabel: r.fixtureLabel || '', quantity: parseNonNegativeNumber(r.quantity, 0), hotCold: !!r.hotCold, unitWsfu: parseNonNegativeNumber(r.unitWsfu, 0), rowTotalWsfu: parseNonNegativeNumber(r.rowTotalWsfu, 0) }; }), results: wsfuZones[i].results || {} });
      }
      return zones;
    }

    function getWsfuFixtureRowsState() {
      var zones = getWsfuZonesState();
      var rows = [];
      for (var i = 0; i < zones.length; i++) rows = rows.concat(zones[i].fixtures || []);
      return rows;
    }

    function updateWsfuTotalsAndUI() {
      var usageType = document.getElementById('wsfuUsageType') ? document.getElementById('wsfuUsageType').value : 'PRIVATE';
      var autoTotal = 0;
      ensureWsfuZonesInitialized();
      for (var i = 0; i < wsfuZones.length; i++) {
        var zone = wsfuZones[i], zoneTotal = 0;
        for (var j = 0; j < zone.fixtures.length; j++) {
          var row = zone.fixtures[j];
          var item = getWsfuFixtureByKey(row.fixtureKey);
          var unit = wsfuFixtureUnitByUsage(item, usageType);
          var factor = row.hotCold ? 0.75 : 1.0;
          row.unitWsfu = unit;
          row.rowTotalWsfu = roundNumber(parseNonNegativeNumber(row.quantity, 0) * unit * factor, 3);
          row.fixtureLabel = item ? item.label : '';
          zoneTotal += row.rowTotalWsfu;
        }
        zone.results = zone.results || {};
        zone.results.zoneTotalFu = roundNumber(zoneTotal, 3);
        autoTotal += zoneTotal;
      }
      autoTotal = roundNumber(autoTotal, 3);
      if (document.getElementById('wsfuAutoTotalFu')) document.getElementById('wsfuAutoTotalFu').value = autoTotal;
      var useManual = !!(document.getElementById('wsfuUseManualTotal') && document.getElementById('wsfuUseManualTotal').checked);
      var manualNode = document.getElementById('wsfuManualTotalFu');
      if (manualNode) manualNode.disabled = !useManual;
      var effective = useManual ? parseNonNegativeNumber(manualNode ? manualNode.value : 0, 0) : autoTotal;
      if (document.getElementById('wsfuTotalFu')) document.getElementById('wsfuTotalFu').value = roundNumber(effective, 3);
      renderWsfuZones();
      if (document.getElementById('wsfuPreview')) document.getElementById('wsfuPreview').innerHTML = 'Zones: ' + wsfuZones.length + ' | Effective WSFU: ' + formatNumber(effective, 2) + '.';
    }

    function renderWsfuResult(response) {
      if (response && response.zones && response.zones.length) {
        var html = '<div class="wsfu-results-section"><div class="wsfu-results-subtitle">Results Summary</div><div class="wsfu-result-grid"><div class="wsfu-result-item"><span class="wsfu-result-label">Effective WSFU</span><span class="wsfu-result-value">' + formatNumber(parseNonNegativeNumber(response.totalFu, 0), 2) + '</span></div><div class="wsfu-result-item"><span class="wsfu-result-label">Service Size</span><span class="wsfu-result-value">' + (response.serviceSize || '') + '</span></div><div class="wsfu-result-item"><span class="wsfu-result-label">Zones</span><span class="wsfu-result-value">' + response.zones.length + '</span></div></div></div>';
        for (var i = 0; i < response.zones.length; i++) {
          var z = response.zones[i];
          html += '<div class="cond-zone-card"><div class="cond-zone-header"><div class="cond-zone-title">' + (z.name || ('Zone ' + (i + 1))) + '</div><div class="module-note">Zone WSFU: ' + formatNumber(parseNonNegativeNumber(z.totalFu, 0), 2) + ' | Size: ' + (z.serviceSize || 'N/A') + '</div></div></div>';
        }
        return html;
      }
      return '<div class="result-placeholder">Enter fixture rows and press Calculate.</div>';
    }

    async function calculateWsfu() {
      updateWsfuTotalsAndUI();
      var flushType = (document.getElementById('wsfuFlushType') && document.getElementById('wsfuFlushType').value === 'FLUSH_VALVE') ? 'FLUSH_VALVE' : 'FLUSH_TANK';
      var designLength = parseFloat(document.getElementById('wsfuDesignLength') ? document.getElementById('wsfuDesignLength').value : '');
      var totalFu = parseFloat(document.getElementById('wsfuTotalFu') ? document.getElementById('wsfuTotalFu').value : '');
      var resultNode = document.getElementById('wsfuResults');
      if (isNaN(totalFu) || totalFu < 0) return setWarningResult(resultNode, 'Enter a non-negative total fixture unit value.');
      if (isNaN(designLength) || designLength < 0) return setWarningResult(resultNode, 'Enter a non-negative design length.');
      if (!window.pywebview || !window.pywebview.api || !window.pywebview.api.calculate_wsfu) return setWarningResult(resultNode, 'WSFU service is not ready yet. Please try again.');
      try {
        var overall = await window.pywebview.api.calculate_wsfu({ totalFu: totalFu, flushType: flushType, designLengthFt: designLength });
        if (!overall || !overall.ok) return setWarningResult(resultNode, (overall && overall.reason ? overall.reason : 'WSFU lookup failed.'));
        var zones = getWsfuZonesState();
        var zoneResults = [];
        for (var i = 0; i < zones.length; i++) {
          var zoneFu = 0;
          for (var j = 0; j < zones[i].fixtures.length; j++) zoneFu += parseNonNegativeNumber(zones[i].fixtures[j].rowTotalWsfu, 0);
          var zr = await window.pywebview.api.calculate_wsfu({ totalFu: zoneFu, flushType: flushType, designLengthFt: designLength });
          zoneResults.push({ id: zones[i].id, name: zones[i].name, totalFu: roundNumber(zoneFu, 3), serviceSize: (zr && zr.ok) ? (zr.serviceSize || '') : '', warning: (zr && !zr.ok) ? (zr.reason || '') : '', fixtureRows: zones[i].fixtures });
        }
        overall.zones = zoneResults;
        latestWsfuSnapshot = { section: 'wsfu', selection: buildProjectPayload().wsfu.selection, result: overall };
        if (resultNode) resultNode.innerHTML = renderWsfuResult(overall);
      } catch (e) {
        latestWsfuSnapshot = null;
        setWarningResult(resultNode, 'WSFU lookup failed. ' + ((e && e.message) ? e.message : 'Unknown error'));
      }
    }

    function resetWsfuSection() {
      latestWsfuSnapshot = null;
      if (document.getElementById('wsfuFlushType')) document.getElementById('wsfuFlushType').value = 'FLUSH_TANK';
      if (document.getElementById('wsfuUsageType')) document.getElementById('wsfuUsageType').value = 'PRIVATE';
      if (document.getElementById('wsfuDesignLength')) document.getElementById('wsfuDesignLength').value = APP_DEFAULTS.wsfuDesignLengthFt || 100;
      if (document.getElementById('wsfuUseManualTotal')) document.getElementById('wsfuUseManualTotal').checked = false;
      if (document.getElementById('wsfuManualTotalFu')) document.getElementById('wsfuManualTotalFu').value = 0;
      wsfuZones = [createDefaultWsfuZone(1)];
      renderWsfuZones();
      updateWsfuTotalsAndUI();
      if (document.getElementById('wsfuResults')) document.getElementById('wsfuResults').innerHTML = 'Enter fixture rows and press Calculate.';
    }

    function makeFixtureUnitRowId() {
      fixtureUnitRowIdSeq += 1;
      return 'fixture_unit_row_' + fixtureUnitRowIdSeq;
    }

    function makeFixtureUnitZoneId() {
      fixtureUnitZoneIdSeq += 1;
      return 'fixture_unit_zone_' + fixtureUnitZoneIdSeq;
    }

    function createDefaultFixtureUnitRow() {
      return {
        id: makeFixtureUnitRowId(),
        fixtureKey: (fixtureUnitCatalog.length > 0 ? fixtureUnitCatalog[0].key : ''),
        category: 'PRIVATE',
        quantity: 1,
        hotCold: false,
        note: '',
        wasteUnit: 0,
        waterUnit: 0,
        rowWasteTotal: 0,
        rowWaterTotal: 0
      };
    }

    function createDefaultFixtureUnitZone(indexHint) {
      var index = parseInt(indexHint, 10);
      if (isNaN(index) || index < 1) index = (fixtureUnitZones && fixtureUnitZones.length ? fixtureUnitZones.length + 1 : 1);
      return {
        id: makeFixtureUnitZoneId(),
        name: 'Zone ' + index,
        rows: [createDefaultFixtureUnitRow()],
        results: {}
      };
    }

    function ensureFixtureUnitZonesInitialized() {
      if (!fixtureUnitZones || fixtureUnitZones.length <= 0) fixtureUnitZones = [createDefaultFixtureUnitZone(1)];
    }

    function getFixtureUnitItemByKey(key) {
      var normalized = normalizeCpcFixtureKey(key);
      for (var i = 0; i < fixtureUnitCatalog.length; i++) {
        if (fixtureUnitCatalog[i].key === normalized || fixtureUnitCatalog[i].key === key) return fixtureUnitCatalog[i];
      }
      return null;
    }

    function fixtureUnitCategoryKey(raw) {
      var value = String(raw || 'PRIVATE').toUpperCase();
      if (value !== 'PUBLIC' && value !== 'ASSEMBLY') value = 'PRIVATE';
      return value;
    }

    function fixtureUnitCategoryLabel(raw) {
      var value = fixtureUnitCategoryKey(raw);
      if (value === 'ASSEMBLY') return 'Assembly';
      if (value === 'PUBLIC') return 'Public';
      return 'Private';
    }

    function fixtureUnitLookupValue(item, side, category) {
      if (!item || !item[side] || typeof item[side] !== 'object') return null;
      var cat = fixtureUnitCategoryKey(category).toLowerCase();
      if (!Object.prototype.hasOwnProperty.call(item[side], cat)) return null;
      var value = item[side][cat];
      if (value === null || typeof value === 'undefined') return null;
      return parseNonNegativeNumber(value, 0);
    }

    function addFixtureUnitZone(seedZone) {
      var zone = createDefaultFixtureUnitZone((fixtureUnitZones ? fixtureUnitZones.length : 0) + 1);
      var i, rows, rowSeed;
      if (seedZone && seedZone.id) zone.id = String(seedZone.id);
      if (seedZone && seedZone.name) zone.name = String(seedZone.name);
      zone.rows = [];
      rows = (seedZone && seedZone.rows && seedZone.rows.length) ? seedZone.rows : [createDefaultFixtureUnitRow()];
      for (i = 0; i < rows.length; i++) {
        rowSeed = rows[i] || {};
        zone.rows.push({
          id: rowSeed.id ? String(rowSeed.id) : makeFixtureUnitRowId(),
          fixtureKey: rowSeed.fixtureKey ? normalizeCpcFixtureKey(rowSeed.fixtureKey) : (fixtureUnitCatalog.length > 0 ? fixtureUnitCatalog[0].key : ''),
          category: fixtureUnitCategoryKey(rowSeed.category || 'PRIVATE'),
          quantity: parseNonNegativeNumber(rowSeed.quantity, 1),
          hotCold: !!rowSeed.hotCold,
          note: rowSeed.note ? String(rowSeed.note) : '',
          wasteUnit: parseNonNegativeNumber((typeof rowSeed.wasteUnit !== 'undefined' ? rowSeed.wasteUnit : rowSeed.sanitaryUnit), 0),
          waterUnit: parseNonNegativeNumber(rowSeed.waterUnit, 0),
          rowWasteTotal: parseNonNegativeNumber((typeof rowSeed.rowWasteTotal !== 'undefined' ? rowSeed.rowWasteTotal : rowSeed.rowSanitaryTotal), 0),
          rowWaterTotal: parseNonNegativeNumber(rowSeed.rowWaterTotal, 0)
        });
      }
      if (zone.rows.length <= 0) zone.rows = [createDefaultFixtureUnitRow()];
      fixtureUnitZones.push(zone);
      renderFixtureUnitRows();
      updateFixtureUnitRowsAndPreview();
    }

    function removeFixtureUnitZone(zoneId) {
      ensureFixtureUnitZonesInitialized();
      if (fixtureUnitZones.length <= 1) fixtureUnitZones = [createDefaultFixtureUnitZone(1)];
      else fixtureUnitZones = fixtureUnitZones.filter(function (zone) { return zone.id !== zoneId; });
      renderFixtureUnitRows();
      updateFixtureUnitRowsAndPreview();
    }

    function addFixtureUnitRow(zoneId, seed) {
      ensureFixtureUnitZonesInitialized();
      var zone = null;
      for (var z = 0; z < fixtureUnitZones.length; z++) if (fixtureUnitZones[z].id === zoneId) zone = fixtureUnitZones[z];
      if (!zone) zone = fixtureUnitZones[0];
      var base = createDefaultFixtureUnitRow();
      if (seed && typeof seed === 'object') {
        base.id = seed.id ? String(seed.id) : base.id;
        base.fixtureKey = seed.fixtureKey ? normalizeCpcFixtureKey(seed.fixtureKey) : base.fixtureKey;
        base.category = fixtureUnitCategoryKey(seed.category || 'PRIVATE');
        base.quantity = parseNonNegativeNumber(seed.quantity, 1);
        base.hotCold = !!seed.hotCold;
        base.note = seed.note ? String(seed.note) : '';
      }
      zone.rows.push(base);
      renderFixtureUnitRows();
      updateFixtureUnitRowsAndPreview();
    }

    function removeFixtureUnitRow(zoneId, rowId) {
      ensureFixtureUnitZonesInitialized();
      var zone = null;
      for (var z = 0; z < fixtureUnitZones.length; z++) if (fixtureUnitZones[z].id === zoneId) zone = fixtureUnitZones[z];
      if (!zone) return;
      if (zone.rows.length <= 1) return;
      zone.rows = zone.rows.filter(function (row) { return row.id !== rowId; });
      renderFixtureUnitRows();
      updateFixtureUnitRowsAndPreview();
    }

    function onFixtureUnitZoneFieldChange(zoneId, value) {
      for (var i = 0; i < fixtureUnitZones.length; i++) {
        if (fixtureUnitZones[i].id !== zoneId) continue;
        fixtureUnitZones[i].name = String(value || '');
      }
      updateFixtureUnitRowsAndPreview();
    }

    function onFixtureUnitRowFieldChange(zoneId, rowId, field, value) {
      ensureFixtureUnitZonesInitialized();
      for (var z = 0; z < fixtureUnitZones.length; z++) {
        if (fixtureUnitZones[z].id !== zoneId) continue;
        for (var i = 0; i < fixtureUnitZones[z].rows.length; i++) {
          if (fixtureUnitZones[z].rows[i].id !== rowId) continue;
          if (field === 'fixtureKey') fixtureUnitZones[z].rows[i].fixtureKey = normalizeCpcFixtureKey(value || fixtureUnitZones[z].rows[i].fixtureKey);
          if (field === 'category') fixtureUnitZones[z].rows[i].category = fixtureUnitCategoryKey(value);
          if (field === 'quantity') fixtureUnitZones[z].rows[i].quantity = parseNonNegativeNumber(value, 0);
          if (field === 'hotCold') fixtureUnitZones[z].rows[i].hotCold = !!value;
          if (field === 'note') fixtureUnitZones[z].rows[i].note = String(value || '');
          break;
        }
      }
      updateFixtureUnitRowsAndPreview();
    }

    function getFixtureUnitSelectionState() {
      return {
        codeBasis: (document.getElementById('fixtureUnitCodeBasis') ? document.getElementById('fixtureUnitCodeBasis').value : 'IPC'),
        drainageOrientation: (document.getElementById('fixtureUnitDrainageOrientation') && document.getElementById('fixtureUnitDrainageOrientation').value === 'VERTICAL') ? 'VERTICAL' : 'HORIZONTAL',
        flushType: (document.getElementById('fixtureUnitFlushType') && document.getElementById('fixtureUnitFlushType').value === 'FLUSH_VALVE') ? 'FLUSH_VALVE' : 'FLUSH_TANK',
        designLengthFt: parseNonNegativeNumber(document.getElementById('fixtureUnitDesignLength') ? document.getElementById('fixtureUnitDesignLength').value : (APP_DEFAULTS.wsfuDesignLengthFt || 100), APP_DEFAULTS.wsfuDesignLengthFt || 100),
        zones: getFixtureUnitZonesState(),
        rows: getFixtureUnitRowsFlat()
      };
    }

    function getFixtureUnitZonesState() {
      ensureFixtureUnitZonesInitialized();
      return fixtureUnitZones.map(function (zone, zoneIndex) {
        return {
          id: zone.id,
          name: zone.name || ('Zone ' + (zoneIndex + 1)),
          rows: (zone.rows || []).map(function (row) {
            var item = getFixtureUnitItemByKey(row.fixtureKey);
            return {
              id: row.id,
              fixtureKey: normalizeCpcFixtureKey(row.fixtureKey),
              fixtureLabel: item ? item.label : '',
              category: fixtureUnitCategoryKey(row.category || 'PRIVATE'),
              quantity: parseNonNegativeNumber(row.quantity, 0),
              hotCold: !!row.hotCold,
              note: row.note || '',
              wasteUnit: parseNonNegativeNumber(row.wasteUnit, 0),
              waterUnit: parseNonNegativeNumber(row.waterUnit, 0),
              rowWasteTotal: parseNonNegativeNumber(row.rowWasteTotal, 0),
              rowWaterTotal: parseNonNegativeNumber(row.rowWaterTotal, 0)
            };
          }),
          results: zone.results || {}
        };
      });
    }

    function getFixtureUnitRowsFlat() {
      var zones = getFixtureUnitZonesState();
      var rows = [];
      for (var i = 0; i < zones.length; i++) rows = rows.concat(zones[i].rows || []);
      return rows;
    }

    function computeFixtureUnitRows() {
      ensureFixtureUnitZonesInitialized();
      var sanitaryTotal = 0;
      var waterTotal = 0;
      for (var z = 0; z < fixtureUnitZones.length; z++) {
        var zone = fixtureUnitZones[z];
        var zoneWaste = 0;
        var zoneWater = 0;
        for (var i = 0; i < zone.rows.length; i++) {
          var row = zone.rows[i];
          var item = getFixtureUnitItemByKey(row.fixtureKey);
          var category = fixtureUnitCategoryKey(row.category || 'PRIVATE');
          var wasteUnitRaw = fixtureUnitLookupValue(item, 'waste', category);
          var waterUnitRaw = fixtureUnitLookupValue(item, 'water', category);
          var waterFactor = row.hotCold ? 0.75 : 1.0;
          row.wasteUnit = (wasteUnitRaw === null) ? 0 : parseNonNegativeNumber(wasteUnitRaw, 0);
          row.waterUnit = (waterUnitRaw === null) ? 0 : roundNumber(parseNonNegativeNumber(waterUnitRaw, 0) * waterFactor, 3);
          row.rowWasteTotal = roundNumber(parseNonNegativeNumber(row.quantity, 0) * row.wasteUnit, 3);
          row.rowWaterTotal = roundNumber(parseNonNegativeNumber(row.quantity, 0) * row.waterUnit, 3);
          zoneWaste += row.rowWasteTotal;
          zoneWater += row.rowWaterTotal;
        }
        zone.results = zone.results || {};
        zone.results.zoneWasteTotal = roundNumber(zoneWaste, 3);
        zone.results.zoneWaterTotal = roundNumber(zoneWater, 3);
        sanitaryTotal += zoneWaste;
        waterTotal += zoneWater;
      }
      return {
        zones: fixtureUnitZones,
        sanitaryTotal: roundNumber(sanitaryTotal, 3),
        waterTotal: roundNumber(waterTotal, 3)
      };
    }

    function renderFixtureUnitRows() {
      var container = document.getElementById('fixtureUnitZonesContainer');
      if (!container) return;
      var html = '';
      var computed = computeFixtureUnitRows();
      var zones = computed.zones || [];
      if (zones.length <= 0) {
        html = '<div class="fixture-unit-empty">No fixture rows yet. Add a zone to begin.</div>';
      } else {
        for (var z = 0; z < zones.length; z++) {
          var zone = zones[z];
          var zoneRows = zone.rows || [];
          html += '<div class="cond-zone-card">';
          html += '<div class="cond-zone-header"><div class="cond-zone-title">Zone ' + (z + 1) + '</div><div class="cond-zone-name-wrap"><span class="cond-zone-name-label">Zone Name</span><input type="text" value="' + (zone.name || ('Zone ' + (z + 1))) + '" oninput="onFixtureUnitZoneFieldChange(\'' + zone.id + '\', this.value);" /></div><div class="zone-subtitle">Waste: ' + formatNumber(parseNonNegativeNumber((zone.results || {}).zoneWasteTotal, 0), 2) + ' | Water: ' + formatNumber(parseNonNegativeNumber((zone.results || {}).zoneWaterTotal, 0), 2) + '</div></div>';
          html += '<div class="cond-zone-body"><div class="fixture-unit-table-wrap"><table class="fixture-unit-table"><thead><tr><th>Fixture Type</th><th>Category</th><th class="num">Qty</th><th class="center">Hot+Cold</th><th>Note</th><th class="num">Waste / Fixture</th><th class="num">Water / Fixture</th><th class="num">Row Waste</th><th class="num">Row Water</th><th class="num">Action</th></tr></thead><tbody>';
          for (var i = 0; i < zoneRows.length; i++) {
            var row = zoneRows[i];
            var options = '';
            for (var j = 0; j < fixtureUnitCatalog.length; j++) {
              var item = fixtureUnitCatalog[j];
              options += '<option value="' + item.key + '"' + (item.key === normalizeCpcFixtureKey(row.fixtureKey) ? ' selected' : '') + '>' + item.label + '</option>';
            }
            html += '<tr>' +
              '<td class="fixture-unit-col-fixture"><select onchange="onFixtureUnitRowFieldChange(\'' + zone.id + '\', \'' + row.id + '\', \'fixtureKey\', this.value);">' + options + '</select></td>' +
              '<td><select onchange="onFixtureUnitRowFieldChange(\'' + zone.id + '\', \'' + row.id + '\', \'category\', this.value);"><option value="PRIVATE"' + (row.category === 'PRIVATE' ? ' selected' : '') + '>Private</option><option value="PUBLIC"' + (row.category === 'PUBLIC' ? ' selected' : '') + '>Public</option><option value="ASSEMBLY"' + (row.category === 'ASSEMBLY' ? ' selected' : '') + '>Assembly</option></select></td>' +
              '<td class="num"><input class="fixture-unit-qty-input" type="number" min="0" step="1" value="' + formatNumber(parseNonNegativeNumber(row.quantity, 0), 0) + '" oninput="onFixtureUnitRowFieldChange(\'' + zone.id + '\', \'' + row.id + '\', \'quantity\', this.value);" /></td>' +
              '<td class="center"><input type="checkbox" class="fixture-unit-hotcold-checkbox" ' + (row.hotCold ? 'checked' : '') + ' onchange="onFixtureUnitRowFieldChange(\'' + zone.id + '\', \'' + row.id + '\', \'hotCold\', this.checked);" /></td>' +
              '<td><input class="fixture-unit-note-input" type="text" value="' + (row.note || '') + '" placeholder="Optional note" oninput="onFixtureUnitRowFieldChange(\'' + zone.id + '\', \'' + row.id + '\', \'note\', this.value);" /></td>' +
              '<td class="num"><span class="fixture-unit-value-badge calculated">' + (parseNonNegativeNumber(row.wasteUnit, 0) > 0 ? formatNumber(parseNonNegativeNumber(row.wasteUnit, 0), 2) : 'N/A') + '</span></td>' +
              '<td class="num"><span class="fixture-unit-value-badge calculated">' + (parseNonNegativeNumber(row.waterUnit, 0) > 0 ? formatNumber(parseNonNegativeNumber(row.waterUnit, 0), 2) : 'N/A') + '</span></td>' +
              '<td class="num"><span class="fixture-unit-value-badge calculated total">' + formatNumber(parseNonNegativeNumber(row.rowWasteTotal, 0), 2) + '</span></td>' +
              '<td class="num"><span class="fixture-unit-value-badge calculated total">' + formatNumber(parseNonNegativeNumber(row.rowWaterTotal, 0), 2) + '</span></td>' +
              '<td class="num"><input type="button" class="btn fixture-unit-remove-btn" value="Remove" onclick="removeFixtureUnitRow(\'' + zone.id + '\', \'' + row.id + '\');" /></td>' +
              '</tr>';
          }
          html += '</tbody></table></div></div>';
          html += '<div class="cond-zone-actions"><div class="cond-zone-actions-left"><input type="button" class="btn wsfu-add-fixture-btn" value="+ ADD FIXTURE" onclick="addFixtureUnitRow(\'' + zone.id + '\');" /><input type="button" class="btn wsfu-add-fixture-btn fixture-zone-add-zone-btn" value="+ ADD ZONE" onclick="addFixtureUnitZone();" /></div><div class="cond-zone-actions-right"><input type="button" class="btn btn-small zone-danger-btn fixture-zone-remove-btn" value="Remove Zone" onclick="removeFixtureUnitZone(\'' + zone.id + '\');" /></div></div>';
          html += '</div>';
        }
      }
      container.innerHTML = html;
    }

    function clearFixtureUnitResults() {
      var sanitaryNode = document.getElementById('fixtureUnitSanitaryResults');
      var waterNode = document.getElementById('fixtureUnitWaterResults');
      if (sanitaryNode) sanitaryNode.innerHTML = '<div class="fixture-unit-result-empty">Add fixtures to generate sanitary drainage results.</div>';
      if (waterNode) waterNode.innerHTML = '<div class="fixture-unit-result-empty">Add fixtures to generate water supply fixture unit results.</div>';
    }

    function renderFixtureUnitResults(snapshot) {
      var sanitaryNode = document.getElementById('fixtureUnitSanitaryResults');
      var waterNode = document.getElementById('fixtureUnitWaterResults');
      if (!snapshot || !sanitaryNode || !waterNode) return;
      var sanitary = snapshot.sanitary || {};
      var water = snapshot.water || {};
      var zones = snapshot.zones || [];
      sanitaryNode.innerHTML =
        '<div class="fixture-unit-result-primary-grid">' +
        '<div class="fixture-unit-result-primary"><span class="fixture-unit-result-primary-label">Total Sanitary Load</span><span class="fixture-unit-result-primary-value">' + formatNumber(parseNonNegativeNumber(sanitary.total, 0), 2) + '</span></div>' +
        '<div class="fixture-unit-result-primary"><span class="fixture-unit-result-primary-label">Suggested Pipe Size</span><span class="fixture-unit-result-primary-value">' + (parseNonNegativeNumber(sanitary.total, 0) > 0 ? (sanitary.recommendedSize || 'Out of range') : 'Pending') + '</span></div>' +
        '</div>' +
        '<div class="fixture-unit-summary-grid compact">' +
        '<div class="fixture-unit-summary-item"><span class="fixture-unit-summary-label">Code Basis</span><span class="fixture-unit-summary-value">' + (sanitary.codeBasis || 'IPC') + '</span></div>' +
        '<div class="fixture-unit-summary-item"><span class="fixture-unit-summary-label">Orientation</span><span class="fixture-unit-summary-value">' + ((sanitary.drainageOrientation === 'VERTICAL') ? 'Vertical' : 'Horizontal') + '</span></div>' +
        '<div class="fixture-unit-summary-item"><span class="fixture-unit-summary-label">Zones</span><span class="fixture-unit-summary-value">' + zones.length + '</span></div>' +
        '</div>' +
        (sanitary.warning ? '<div class="warning-line">' + sanitary.warning + '</div>' : '<div class="fixture-unit-note-line">' + (parseNonNegativeNumber(sanitary.total, 0) > 0 ? 'Drainage fixture units are based on the selected code basis.' : 'Enter fixture quantities greater than zero to evaluate sanitary sizing.') + '</div>');

      waterNode.innerHTML =
        '<div class="fixture-unit-result-primary-grid">' +
        '<div class="fixture-unit-result-primary"><span class="fixture-unit-result-primary-label">Total WSFU</span><span class="fixture-unit-result-primary-value">' + formatNumber(parseNonNegativeNumber(water.total, 0), 2) + '</span></div>' +
        '<div class="fixture-unit-result-primary"><span class="fixture-unit-result-primary-label">Estimated Demand</span><span class="fixture-unit-result-primary-value">' + formatNumber(parseNonNegativeNumber(water.selectedFlowGpm, 0), 2) + ' GPM</span></div>' +
        '<div class="fixture-unit-result-primary"><span class="fixture-unit-result-primary-label">Suggested Pipe Size</span><span class="fixture-unit-result-primary-value">' + (water.serviceSize || 'Pending') + '</span></div>' +
        '</div>' +
        '<div class="fixture-unit-summary-grid compact">' +
        '<div class="fixture-unit-summary-item"><span class="fixture-unit-summary-label">Sizing Basis</span><span class="fixture-unit-summary-value">' + ((water.flushType === 'FLUSH_VALVE') ? 'Flush Valve' : 'Flush Tank') + '</span></div>' +
        '<div class="fixture-unit-summary-item"><span class="fixture-unit-summary-label">Zones</span><span class="fixture-unit-summary-value">' + zones.length + '</span></div>' +
        '</div>' +
        (water.warning ? '<div class="warning-line">' + water.warning + '</div>' : '<div class="fixture-unit-note-line">' + (water.source || 'Water supply sizing follows the WSFU service workflow.') + '</div>');
    }

    async function calculateFixtureUnit(commitResults) {
      if (commitResults !== false) commitResults = true;
      var resultNodeSanitary = document.getElementById('fixtureUnitSanitaryResults');
      var resultNodeWater = document.getElementById('fixtureUnitWaterResults');
      var selection = getFixtureUnitSelectionState();
      var computed = computeFixtureUnitRows();
      var sanitaryMatch = findSanitary7032ByDfu(computed.sanitaryTotal, selection.drainageOrientation);
      var result = {
        sanitary: {
          total: computed.sanitaryTotal,
          recommendedSize: sanitaryMatch ? sanitaryMatch.size : '',
          warning: sanitaryMatch ? '' : (computed.sanitaryTotal > 0 ? 'Sanitary fixture load exceeds Table 703.2 range for selected orientation.' : ''),
          codeBasis: selection.codeBasis,
          drainageOrientation: selection.drainageOrientation
        },
        water: {
          total: computed.waterTotal,
          flushType: selection.flushType,
          designLengthFt: selection.designLengthFt,
          selectedFlowGpm: 0,
          serviceSize: '',
          warning: '',
          source: ''
        },
        fixtureRows: selection.rows,
        zones: selection.zones,
        timestamp: (new Date()).toISOString()
      };

      if (computed.waterTotal > 0) {
        if (!window.pywebview || !window.pywebview.api || !window.pywebview.api.calculate_wsfu) {
          result.water.warning = 'WSFU service is not ready yet. Please try again.';
        } else {
          try {
            var waterResult = await window.pywebview.api.calculate_wsfu({
              totalFu: computed.waterTotal,
              flushType: selection.flushType,
              designLengthFt: selection.designLengthFt
            });
            if (waterResult && waterResult.ok) {
              result.water.selectedFlowGpm = parseNonNegativeNumber(waterResult.selectedFlowGpm, 0);
              result.water.serviceSize = waterResult.serviceSize || '';
              result.water.warning = waterResult.warning || '';
              result.water.source = waterResult.source || '';
            } else {
              result.water.warning = (waterResult && waterResult.reason) ? waterResult.reason : 'WSFU sizing lookup failed.';
            }
          } catch (e) {
            result.water.warning = 'WSFU sizing lookup failed. ' + ((e && e.message) ? e.message : 'Unknown error');
          }
        }
      } else {
        result.water.warning = 'Enter fixture quantities greater than zero to evaluate water supply sizing.';
      }

      latestFixtureUnitSnapshot = {
        section: 'fixtureUnit',
        selection: selection,
        result: result
      };
      if (commitResults || resultNodeSanitary || resultNodeWater) renderFixtureUnitResults(result);
      return result;
    }

    function updateFixtureUnitRowsAndPreview() {
      renderFixtureUnitRows();
      var previewNode = document.getElementById('fixtureUnitPreview');
      var summaryNode = document.getElementById('fixtureUnitSummaryStrip');
      var computed = computeFixtureUnitRows();
      var rowCount = 0;
      for (var i = 0; i < fixtureUnitZones.length; i++) rowCount += (fixtureUnitZones[i].rows || []).length;
      if (summaryNode) {
        summaryNode.innerHTML =
          '<span class="fixture-unit-summary-pill"><span class="k">Total Zones</span><span class="v">' + fixtureUnitZones.length + '</span></span>' +
          '<span class="fixture-unit-summary-pill"><span class="k">Total Fixture Rows</span><span class="v">' + rowCount + '</span></span>' +
          '<span class="fixture-unit-summary-pill"><span class="k">Total Waste</span><span class="v">' + formatNumber(computed.sanitaryTotal, 2) + '</span></span>' +
          '<span class="fixture-unit-summary-pill"><span class="k">Total Water</span><span class="v">' + formatNumber(computed.waterTotal, 2) + '</span></span>';
      }
      if (previewNode) previewNode.innerHTML = '';
      if (fixtureUnitCalcTimer) {
        window.clearTimeout(fixtureUnitCalcTimer);
        fixtureUnitCalcTimer = null;
      }
      fixtureUnitCalcTimer = window.setTimeout(function () {
        fixtureUnitCalcTimer = null;
        calculateFixtureUnit(false);
      }, 180);
    }

    function resetFixtureUnitSection() {
      latestFixtureUnitSnapshot = null;
      if (fixtureUnitCalcTimer) {
        window.clearTimeout(fixtureUnitCalcTimer);
        fixtureUnitCalcTimer = null;
      }
      if (document.getElementById('fixtureUnitCodeBasis')) document.getElementById('fixtureUnitCodeBasis').value = 'IPC';
      if (document.getElementById('fixtureUnitDrainageOrientation')) document.getElementById('fixtureUnitDrainageOrientation').value = 'HORIZONTAL';
      if (document.getElementById('fixtureUnitFlushType')) document.getElementById('fixtureUnitFlushType').value = 'FLUSH_TANK';
      if (document.getElementById('fixtureUnitDesignLength')) document.getElementById('fixtureUnitDesignLength').value = APP_DEFAULTS.wsfuDesignLengthFt || 100;
      fixtureUnitZones = [createDefaultFixtureUnitZone(1)];
      renderFixtureUnitRows();
      clearFixtureUnitResults();
      updateFixtureUnitRowsAndPreview();
    }

    function addDuctZone(seedZone) {
      var zone = createDefaultDuctZone(ductZones.length + 1);
      var segs, i, seg;
      if (seedZone && seedZone.id) zone.id = String(seedZone.id);
      if (seedZone && seedZone.name) zone.name = String(seedZone.name);
      zone.segments = [];
      segs = (seedZone && seedZone.segments && seedZone.segments.length) ? seedZone.segments : [createDefaultDuctSegment()];
      for (i = 0; i < segs.length; i++) {
        seg = segs[i] || {};
        zone.segments.push({
          id: seg.id || makeDuctSegmentId(),
          name: seg.name || ('Segment ' + (i + 1)),
          mode: normalizeDuctMode(seg.mode || DUCT_CALC.modes.SIZE_FROM_CFM_FRICTION, seg.shape || DUCT_CALC.shapes.ROUND),
          shape: (seg.shape === DUCT_CALC.shapes.RECTANGULAR) ? DUCT_CALC.shapes.RECTANGULAR : DUCT_CALC.shapes.ROUND,
          cfm: parseNonNegativeNumber(seg.cfm, 0),
          frictionTarget: parseNonNegativeNumber(seg.frictionTarget, 0),
          velocityTarget: parseNonNegativeNumber(seg.velocityTarget, 0),
          diameter: parseNonNegativeNumber(seg.diameter, 0),
          width: parseNonNegativeNumber(seg.width, 0),
          height: parseNonNegativeNumber(seg.height, 0),
          ratioLimit: parseNonNegativeNumber(seg.ratioLimit, 3)
        });
      }
      if (zone.segments.length <= 0) zone.segments.push(createDefaultDuctSegment());
      ductZones.push(zone);
      renderDuctZones();
      calculateDuct(false);
    }

    function removeDuctZone(zoneId) {
      if (ductZones.length <= 1) ductZones = [createDefaultDuctZone(1)];
      else ductZones = ductZones.filter(function (z) { return z.id !== zoneId; });
      renderDuctZones();
      calculateDuct(false);
    }

    function addDuctSegmentRow(zoneId, seed) {
      var zone = ductZones.find(function (z) { return z.id === zoneId; });
      if (!zone) return;
      zone.segments.push(Object.assign(createDefaultDuctSegment(), seed || {}));
      renderDuctZones();
      calculateDuct(false);
    }

    function removeDuctSegmentRow(zoneId, segmentId) {
      var zone = ductZones.find(function (z) { return z.id === zoneId; });
      if (!zone) return;
      if (zone.segments.length <= 1) return;
      zone.segments = zone.segments.filter(function (s) { return s.id !== segmentId; });
      renderDuctZones();
      calculateDuct(false);
    }

    function onDuctZoneFieldChange(zoneId, value) {
      var zone = ductZones.find(function (z) { return z.id === zoneId; });
      if (zone) zone.name = String(value || '');
    }

    function onDuctSegmentFieldChange(zoneId, segmentId, field, value) {
      var zone = ductZones.find(function (z) { return z.id === zoneId; });
      if (!zone) return;
      var seg = zone.segments.find(function (s) { return s.id === segmentId; });
      if (!seg) return;
      if (field === 'mode') seg.mode = normalizeDuctMode(value, seg.shape);
      else if (field === 'shape') seg.shape = (value === DUCT_CALC.shapes.RECTANGULAR ? DUCT_CALC.shapes.RECTANGULAR : DUCT_CALC.shapes.ROUND);
      else if (field === 'name') seg.name = String(value || '');
      else seg[field] = parseNonNegativeNumber(value, 0);
      calculateDuct(false);
    }

    function renderDuctZones() {
      var node = document.getElementById('ductZonesContainer');
      var html = '';
      if (!node) return;
      ensureDuctZonesInitialized();
      for (var i = 0; i < ductZones.length; i++) {
        var zone = ductZones[i];
        var segCount = (zone.segments || []).length;
        var zoneCfm = 0;
        for (var zc = 0; zc < segCount; zc++) zoneCfm += parseNonNegativeNumber((zone.segments[zc] || {}).cfm, 0);
        html += '<div class="cond-zone-card"><div class="cond-zone-header"><div class="cond-zone-title">Zone ' + (i + 1) + '</div><div class="cond-zone-name-wrap"><span class="cond-zone-name-label">Zone Name</span><input type="text" value="' + (zone.name || ('Zone ' + (i + 1))) + '" oninput="onDuctZoneFieldChange(\'' + zone.id + '\', this.value);" /></div><div class="zone-subtitle">' + segCount + ' segments | ' + formatNumber(zoneCfm, 0) + ' CFM</div></div><div class="cond-zone-body">';
        for (var j = 0; j < zone.segments.length; j++) {
          var s = zone.segments[j];
          html += '<div class="duct-segment-card">';
          html += '<div class="duct-segment-grid duct-segment-grid-top">';
          html += '<div class="duct-seg-field"><span class="duct-seg-label">Segment</span><input type="text" value="' + (s.name || ('Segment ' + (j + 1))) + '" oninput="onDuctSegmentFieldChange(\'' + zone.id + '\', \'' + s.id + '\', \'name\', this.value);" /></div>';
          html += '<div class="duct-seg-field"><span class="duct-seg-label">Mode</span><select onchange="onDuctSegmentFieldChange(\'' + zone.id + '\', \'' + s.id + '\', \'mode\', this.value);"><option value="SIZE_FROM_CFM_FRICTION"' + (s.mode === 'SIZE_FROM_CFM_FRICTION' ? ' selected' : '') + '>CFM+Friction</option><option value="SIZE_FROM_CFM_VELOCITY"' + (s.mode === 'SIZE_FROM_CFM_VELOCITY' ? ' selected' : '') + '>CFM+Velocity</option><option value="CHECK_ROUND_SIZE"' + (s.mode === 'CHECK_ROUND_SIZE' ? ' selected' : '') + '>Check Round</option><option value="CHECK_RECT_SIZE"' + (s.mode === 'CHECK_RECT_SIZE' ? ' selected' : '') + '>Check Rect</option></select></div>';
          html += '<div class="duct-seg-field"><span class="duct-seg-label">Shape</span><select onchange="onDuctSegmentFieldChange(\'' + zone.id + '\', \'' + s.id + '\', \'shape\', this.value);"><option value="ROUND"' + (s.shape === 'ROUND' ? ' selected' : '') + '>Round</option><option value="RECTANGULAR"' + (s.shape === 'RECTANGULAR' ? ' selected' : '') + '>Rect</option></select></div>';
          html += '<div class="duct-seg-field"><span class="duct-seg-label">CFM</span><input type="number" step="0.01" value="' + parseNonNegativeNumber(s.cfm, 0) + '" oninput="onDuctSegmentFieldChange(\'' + zone.id + '\', \'' + s.id + '\', \'cfm\', this.value);" /></div>';
          html += '</div>';
          html += '<div class="duct-segment-grid duct-segment-grid-bottom">';
          html += '<div class="duct-seg-field"><span class="duct-seg-label">Friction</span><input type="number" step="0.0001" value="' + parseNonNegativeNumber(s.frictionTarget, 0) + '" oninput="onDuctSegmentFieldChange(\'' + zone.id + '\', \'' + s.id + '\', \'frictionTarget\', this.value);" /></div>';
          html += '<div class="duct-seg-field"><span class="duct-seg-label">Velocity</span><input type="number" step="0.01" value="' + parseNonNegativeNumber(s.velocityTarget, 0) + '" oninput="onDuctSegmentFieldChange(\'' + zone.id + '\', \'' + s.id + '\', \'velocityTarget\', this.value);" /></div>';
          html += '<div class="duct-seg-field duct-seg-output"><span class="duct-seg-label">Dia</span><input type="number" step="0.01" value="' + parseNonNegativeNumber(s.diameter, 0) + '" oninput="onDuctSegmentFieldChange(\'' + zone.id + '\', \'' + s.id + '\', \'diameter\', this.value);" /></div>';
          html += '<div class="duct-seg-field duct-seg-output"><span class="duct-seg-label">W</span><input type="number" step="0.01" value="' + parseNonNegativeNumber(s.width, 0) + '" oninput="onDuctSegmentFieldChange(\'' + zone.id + '\', \'' + s.id + '\', \'width\', this.value);" /></div>';
          html += '<div class="duct-seg-field duct-seg-output"><span class="duct-seg-label">H</span><input type="number" step="0.01" value="' + parseNonNegativeNumber(s.height, 0) + '" oninput="onDuctSegmentFieldChange(\'' + zone.id + '\', \'' + s.id + '\', \'height\', this.value);" /></div>';
          html += '<div class="duct-seg-field duct-seg-action"><span class="duct-seg-label">&nbsp;</span><input type="button" class="btn btn-small zone-danger-btn" value="Remove" onclick="removeDuctSegmentRow(\'' + zone.id + '\', \'' + s.id + '\');" /></div>';
          html += '</div>';
          html += '</div>';
        }
        html += '</div><div class="cond-zone-actions"><div class="cond-zone-actions-left"><input type="button" class="btn cond-secondary-btn" value="+ ADD SEGMENT" onclick="addDuctSegmentRow(\'' + zone.id + '\');" /></div><div class="cond-zone-actions-right"><input type="button" class="btn btn-small zone-danger-btn" value="Remove Zone" onclick="removeDuctZone(\'' + zone.id + '\');" /></div></div></div>';
      }
      node.innerHTML = html;
    }

    function getDuctZonesState() {
      ensureDuctZonesInitialized();
      return ductZones.map(function (z) {
        return { id: z.id, name: z.name, segments: (z.segments || []).map(function (s) { return { id: s.id, name: s.name, mode: s.mode, shape: s.shape, cfm: parseNonNegativeNumber(s.cfm, 0), frictionTarget: parseNonNegativeNumber(s.frictionTarget, 0), velocityTarget: parseNonNegativeNumber(s.velocityTarget, 0), diameter: parseNonNegativeNumber(s.diameter, 0), width: parseNonNegativeNumber(s.width, 0), height: parseNonNegativeNumber(s.height, 0), ratioLimit: parseNonNegativeNumber(s.ratioLimit, 3) }; }), results: z.results || {} };
      });
    }

    function getDuctSelectionState() {
      var zones = getDuctZonesState();
      var first = (zones[0] && zones[0].segments && zones[0].segments[0]) ? zones[0].segments[0] : createDefaultDuctSegment();
      return Object.assign({}, first, { zones: zones });
    }

    function computeDuctSegmentResult(sel) {
      var mode = normalizeDuctMode(sel.mode || DUCT_CALC.modes.SIZE_FROM_CFM_FRICTION, sel.shape || DUCT_CALC.shapes.ROUND);
      var cfm = parseNonNegativeNumber(sel.cfm, 0);
      var eqDiameter = 0, roundSize = '', rectSize = '', velocity = 0, vp = 0, friction = 0, warning = '';
      var velocityTarget = parseNonNegativeNumber(sel.velocityTarget, 0);
      var ratioLimit = parseNonNegativeNumber(sel.ratioLimit, 3);
      var targetRound, chosenRound, chosenRect;
      if (cfm <= 0) return { ok: false, warning: 'Enter airflow (CFM) greater than zero.' };
      if (mode === DUCT_CALC.modes.SIZE_FROM_CFM_FRICTION) {
        if (parseNonNegativeNumber(sel.frictionTarget, 0) <= 0) return { ok: false, warning: 'Enter friction target greater than zero.' };
        targetRound = ductRoundSizeFromTarget(cfm, parseNonNegativeNumber(sel.frictionTarget, 0));
        if (velocityTarget > 0) {
          var vd = Math.sqrt((576 * cfm) / (Math.PI * velocityTarget));
          if (!isNaN(vd) && vd > targetRound) targetRound = vd;
        }
        chosenRound = ductChooseStandardRound(targetRound);
        eqDiameter = chosenRound;
        roundSize = chosenRound + ' in';
        chosenRect = ductChooseRectFromEquivalent(eqDiameter, ratioLimit);
        rectSize = chosenRect ? (chosenRect.width + ' x ' + chosenRect.height + ' in') : '';
        velocity = ductVelocityFpm(cfm, DUCT_CALC.shapes.ROUND, chosenRound, 0, 0);
        vp = ductVelocityPressure(velocity);
        friction = ductFrictionRate(cfm, eqDiameter);
      } else if (mode === DUCT_CALC.modes.SIZE_FROM_CFM_VELOCITY) {
        if (velocityTarget <= 0) return { ok: false, warning: 'Enter velocity target greater than zero.' };
        targetRound = Math.sqrt((576 * cfm) / (Math.PI * velocityTarget));
        chosenRound = ductChooseStandardRound(targetRound);
        eqDiameter = chosenRound;
        roundSize = chosenRound + ' in';
        chosenRect = ductChooseRectFromEquivalent(eqDiameter, ratioLimit);
        rectSize = chosenRect ? (chosenRect.width + ' x ' + chosenRect.height + ' in') : '';
        velocity = ductVelocityFpm(cfm, DUCT_CALC.shapes.ROUND, chosenRound, 0, 0);
        vp = ductVelocityPressure(velocity);
        friction = ductFrictionRate(cfm, eqDiameter);
      } else if (mode === DUCT_CALC.modes.CHECK_ROUND_SIZE) {
        if (parseNonNegativeNumber(sel.diameter, 0) <= 0) return { ok: false, warning: 'Enter round diameter greater than zero.' };
        eqDiameter = parseNonNegativeNumber(sel.diameter, 0);
        roundSize = eqDiameter + ' in';
        chosenRect = ductChooseRectFromEquivalent(eqDiameter, ratioLimit);
        rectSize = chosenRect ? (chosenRect.width + ' x ' + chosenRect.height + ' in') : '';
        velocity = ductVelocityFpm(cfm, DUCT_CALC.shapes.ROUND, eqDiameter, 0, 0);
        vp = ductVelocityPressure(velocity);
        friction = ductFrictionRate(cfm, eqDiameter);
      } else {
        if (parseNonNegativeNumber(sel.width, 0) <= 0 || parseNonNegativeNumber(sel.height, 0) <= 0) return { ok: false, warning: 'Enter rectangular width/height greater than zero.' };
        eqDiameter = ductEquivalentDiameterIn(DUCT_CALC.shapes.RECTANGULAR, 0, parseNonNegativeNumber(sel.width, 0), parseNonNegativeNumber(sel.height, 0));
        chosenRound = ductChooseStandardRound(eqDiameter);
        roundSize = chosenRound + ' in (eq)';
        rectSize = parseNonNegativeNumber(sel.width, 0) + ' x ' + parseNonNegativeNumber(sel.height, 0) + ' in';
        velocity = ductVelocityFpm(cfm, DUCT_CALC.shapes.RECTANGULAR, 0, parseNonNegativeNumber(sel.width, 0), parseNonNegativeNumber(sel.height, 0));
        vp = ductVelocityPressure(velocity);
        friction = ductFrictionRate(cfm, eqDiameter);
      }
      if (velocityTarget > 0 && velocity > velocityTarget) warning = 'Calculated velocity exceeds target.';
      return { ok: true, modeLabel: getDuctModeLabel(mode), equivalentDiameter: roundNumber(eqDiameter, 3), recommendedRound: roundSize, recommendedRect: rectSize, velocityFpm: roundNumber(velocity, 1), velocityPressure: roundNumber(vp, 4), frictionRate: roundNumber(friction, 4), warning: warning };
    }

    function renderDuctResult(result) {
      if (!result) return '<div class="result-placeholder">' + (APP_MESSAGES.ductPlaceholder || 'Enter airflow and sizing criteria to see live duct size recommendations.') + '</div>';
      if (!result.zones) {
        return '<div class="duct-result-summary"><div class="duct-result-metric"><span class="duct-result-metric-label">Equivalent Diameter</span><span class="duct-result-metric-value">' + formatNumber(parseNonNegativeNumber(result.equivalentDiameter, 0), 3) + ' in</span></div><div class="duct-result-metric duct-result-metric-main"><span class="duct-result-metric-label">Recommended Round Size</span><span class="duct-result-metric-value">' + (result.recommendedRound || 'N/A') + '</span></div><div class="duct-result-metric duct-result-metric-main"><span class="duct-result-metric-label">Recommended Rectangular Size</span><span class="duct-result-metric-value">' + (result.recommendedRect || 'N/A') + '</span></div></div>';
      }
      var html = '<div class="duct-result-summary"><div class="duct-result-metric"><span class="duct-result-metric-label">Zones</span><span class="duct-result-metric-value">' + result.zoneCount + '</span></div><div class="duct-result-metric"><span class="duct-result-metric-label">Segments</span><span class="duct-result-metric-value">' + result.segmentCount + '</span></div><div class="duct-result-metric duct-result-metric-main"><span class="duct-result-metric-label">Calculated</span><span class="duct-result-metric-value">' + result.calculatedCount + '</span></div></div>';
      for (var i = 0; i < result.zones.length; i++) {
        var z = result.zones[i];
        html += '<div class="cond-zone-card"><div class="cond-zone-header"><div class="cond-zone-title">' + (z.name || ('Zone ' + (i + 1))) + '</div></div><div class="cond-zone-body"><table class="duct-result-table"><thead><tr><th>Segment</th><th>Mode</th><th class="num">Eq Dia</th><th class="num">Round</th><th class="num">Rect</th><th class="num">Velocity</th><th class="num">Friction</th><th>Warning</th></tr></thead><tbody>';
        for (var j = 0; j < z.rows.length; j++) {
          var r = z.rows[j];
          html += '<tr><td>' + (r.name || ('Segment ' + (j + 1))) + '</td><td>' + (r.modeLabel || '') + '</td><td class="num">' + (r.ok ? formatNumber(r.equivalentDiameter, 3) : '-') + '</td><td class="num">' + (r.ok ? r.recommendedRound : '-') + '</td><td class="num">' + (r.ok ? r.recommendedRect : '-') + '</td><td class="num">' + (r.ok ? (formatNumber(r.velocityFpm, 1) + ' FPM') : '-') + '</td><td class="num">' + (r.ok ? (formatNumber(r.frictionRate, 4) + ' in.wg/100ft') : '-') + '</td><td>' + (r.warning || '') + '</td></tr>';
        }
        html += '</tbody></table></div></div>';
      }
      return html;
    }

    function calculateDuct(commitResults) {
      var zones = getDuctZonesState();
      var resultNode = document.getElementById('ductResults');
      var zoneResults = [];
      var segmentCount = 0;
      var calculatedCount = 0;
      for (var i = 0; i < zones.length; i++) {
        var zr = { id: zones[i].id, name: zones[i].name, rows: [] };
        for (var j = 0; j < zones[i].segments.length; j++) {
          segmentCount += 1;
          var rowResult = computeDuctSegmentResult(zones[i].segments[j]);
          if (rowResult.ok) calculatedCount += 1;
          rowResult.name = zones[i].segments[j].name;
          zr.rows.push(rowResult);
        }
        zoneResults.push(zr);
      }
      var statusText = 'Zones: ' + zones.length + ' | Segments: ' + segmentCount + ' | Calculated: ' + calculatedCount;
      setDuctPreviewContent({}, statusText, calculatedCount === segmentCount ? 'ok' : 'warn');
      latestDuctSnapshot = { section: 'duct', selection: { zones: zones }, result: { zones: zoneResults, zoneCount: zones.length, segmentCount: segmentCount, calculatedCount: calculatedCount } };
      if (resultNode) resultNode.innerHTML = renderDuctResult(latestDuctSnapshot.result);
    }

    function updateDuctModeUI() {
      renderDuctZones();
    }

    function onDuctInputsChanged() {
      calculateDuct(false);
    }

    function resetDuctSection() {
      latestDuctSnapshot = null;
      ductZones = [createDefaultDuctZone(1)];
      renderDuctZones();
      setDuctPreviewContent({}, 'Add duct segment inputs by zone, then click CALCULATE.', '');
      setDuctResultPlaceholder(APP_MESSAGES.ductPlaceholder || 'Enter airflow and sizing criteria to see live duct size recommendations.');
    }

    function makeVentilationZoneId() {
      ventilationZoneIdSeq += 1;
      return 'ventilation_zone_' + ventilationZoneIdSeq;
    }

    function cfmToLs(cfm) {
      return parseNonNegativeNumber(cfm, 0) * 0.47194745;
    }

    function getVentilationOccupancyByKey(key) {
      var target = String(key || '').trim();
      if (!target) return null;
      for (var i = 0; i < ventilationOccupancyCatalog.length; i++) {
        if (ventilationOccupancyCatalog[i].key === target) return ventilationOccupancyCatalog[i];
      }
      return null;
    }

    function resolveVentilationOccupancyKey(savedKey, savedLabel) {
      var key = String(savedKey || '').trim();
      var label = String(savedLabel || '').toLowerCase().trim();
      var i;
      for (i = 0; i < ventilationOccupancyCatalog.length; i++) {
        if (ventilationOccupancyCatalog[i].key === key) return key;
      }
      if (label) {
        for (i = 0; i < ventilationOccupancyCatalog.length; i++) {
          if (String(ventilationOccupancyCatalog[i].occupancy_category || '').toLowerCase().trim() === label) {
            return ventilationOccupancyCatalog[i].key;
          }
        }
      }
      return ventilationOccupancyCatalog.length ? ventilationOccupancyCatalog[0].key : '';
    }

    function ventilationOccupancyMatches(occ, filterText) {
      var f = String(filterText || '').toLowerCase().trim();
      var aliases;
      var i;
      if (!f) return true;
      if (String(occ.occupancy_category || '').toLowerCase().indexOf(f) >= 0) return true;
      if (String(occ.category_group || '').toLowerCase().indexOf(f) >= 0) return true;
      aliases = occ.searchable_aliases || [];
      for (i = 0; i < aliases.length; i++) {
        if (String(aliases[i] || '').toLowerCase().indexOf(f) >= 0) return true;
      }
      return false;
    }

    function createDefaultVentilationZone(index) {
      return {
        id: makeVentilationZoneId(),
        zoneName: 'Zone ' + index,
        spaceName: '',
        occupancyKey: '',
        area: 0,
        populationMode: 'DEFAULT',
        manualPopulation: 0,
        ezPreset: 'DEFAULT_DESIGN_BASIS_08',
        ez: 0.8,
        providedSupply: 0,
        providedSupplySpecified: false,
        notes: ''
      };
    }

    function ensureVentilationInitialized() {
      if (!ventilationZones) ventilationZones = [];
      if (ventilationZoneIdSeq < ventilationZones.length) ventilationZoneIdSeq = ventilationZones.length;
    }

    function renderVentilationOccupancyOptions(filterText) {
      var node = document.getElementById('ventCalcOccupancy');
      var html = '<option value="">Select occupancy category...</option>';
      var selected = node ? String(node.value || '') : '';
      var prevGroup = '';
      if (!node) return;
      for (var i = 0; i < ventilationOccupancyCatalog.length; i++) {
        var row = ventilationOccupancyCatalog[i];
        if (!ventilationOccupancyMatches(row, filterText)) continue;
        if (row.category_group !== prevGroup) {
          html += '<option disabled="disabled">-- ' + row.category_group + ' --</option>';
          prevGroup = row.category_group;
        }
        html += '<option value="' + row.key + '"' + (row.key === selected ? ' selected' : '') + '>[' + row.category_group + '] ' + row.occupancy_category + '</option>';
      }
      node.innerHTML = html;
    }

    function filterVentilationOccupancyOptions(filterText) {
      renderVentilationOccupancyOptions(filterText);
      onVentilationEntryCategoryChanged();
    }

    function setVentilationOccupancySelection(occupancyKey, searchText) {
      var occ = getVentilationOccupancyByKey(occupancyKey);
      var key = occ ? occ.key : String(occupancyKey || '');
      var searchNode = document.getElementById('ventCalcOccupancySearch');
      var selectNode = document.getElementById('ventCalcOccupancy');
      var filterText = (searchText !== undefined && searchText !== null)
        ? String(searchText || '')
        : (occ ? String(occ.occupancy_category || '') : '');
      var found = false;
      var i;
      if (searchNode) searchNode.value = filterText;
      renderVentilationOccupancyOptions(filterText);
      if (selectNode) {
        for (i = 0; i < selectNode.options.length; i++) {
          if (selectNode.options[i].value === key) {
            found = true;
            break;
          }
        }
        if (!found) {
          if (searchNode) searchNode.value = '';
          renderVentilationOccupancyOptions('');
        }
        selectNode.value = key;
      }
    }

    function renderVentilationEzOptions() {
      var node = document.getElementById('ventCalcEzPreset');
      var html = '';
      if (!node) return;
      for (var i = 0; i < ventilationEzCatalog.length; i++) {
        var item = ventilationEzCatalog[i];
        html += '<option value="' + item.key + '">' + item.label + ' (Ez=' + formatNumber(parseNonNegativeNumber(item.ez, 1), 2) + ')</option>';
      }
      node.innerHTML = html;
      if (!node.value) node.value = 'DEFAULT_DESIGN_BASIS_08';
    }

    function getVentilationPresetTokenMap() {
      return {
        office: 'office',
        conference: 'conference',
        classroom: 'classroom',
        lobby: 'lobby',
        corridor: 'corridor',
        restaurant: 'restaurant',
        retail: 'retail',
        gym: 'gym'
      };
    }

    function applyVentilationPreset(name) {
      var tokenMap = getVentilationPresetTokenMap();
      var token = tokenMap[String(name || '').toLowerCase()] || '';
      if (!token) return;
      for (var i = 0; i < ventilationOccupancyCatalog.length; i++) {
        var occ = ventilationOccupancyCatalog[i];
        if (String(occ.occupancy_category || '').toLowerCase().indexOf(token) >= 0) {
          setVentilationOccupancySelection(occ.key, occ.occupancy_category);
          onVentilationEntryCategoryChanged();
          return;
        }
        var aliases = occ.searchable_aliases || [];
        for (var j = 0; j < aliases.length; j++) {
          if (String(aliases[j] || '').toLowerCase().indexOf(token) >= 0) {
            setVentilationOccupancySelection(occ.key, occ.occupancy_category);
            onVentilationEntryCategoryChanged();
            return;
          }
        }
      }
      onVentilationEntryCategoryChanged();
    }

    function getVentilationEntryDraft() {
      var occupancy = getVentilationOccupancyByKey(document.getElementById('ventCalcOccupancy') ? document.getElementById('ventCalcOccupancy').value : '');
      var area = parseNonNegativeNumber(document.getElementById('ventCalcArea') ? document.getElementById('ventCalcArea').value : 0, 0);
      var popMode = (document.getElementById('ventCalcPopulationMode') && document.getElementById('ventCalcPopulationMode').value === 'MANUAL') ? 'MANUAL' : 'DEFAULT';
      var manualPopulation = parseNonNegativeNumber(document.getElementById('ventCalcManualPopulation') ? document.getElementById('ventCalcManualPopulation').value : 0, 0);
      var providedSupplyRaw = document.getElementById('ventCalcProvidedSupply') ? String(document.getElementById('ventCalcProvidedSupply').value || '').trim() : '';
      var providedSupplySpecified = providedSupplyRaw !== '';
      var providedSupply = parseNonNegativeNumber(providedSupplyRaw, 0);
      var pz = popMode === 'MANUAL'
        ? manualPopulation
        : (area * parseNonNegativeNumber(occupancy ? occupancy.default_occupant_density : 0, 0) / 1000);
      return {
        zoneName: document.getElementById('ventCalcZoneName') ? String(document.getElementById('ventCalcZoneName').value || '').trim() : '',
        spaceName: document.getElementById('ventCalcSpaceName') ? String(document.getElementById('ventCalcSpaceName').value || '').trim() : '',
        occupancyKey: occupancy ? occupancy.key : '',
        area: area,
        populationMode: popMode,
        manualPopulation: manualPopulation,
        ezPreset: document.getElementById('ventCalcEzPreset') ? String(document.getElementById('ventCalcEzPreset').value || 'DEFAULT_DESIGN_BASIS_08') : 'DEFAULT_DESIGN_BASIS_08',
        ez: parseNonNegativeNumber(document.getElementById('ventCalcEz') ? document.getElementById('ventCalcEz').value : 0.8, 0.8),
        providedSupply: providedSupply,
        providedSupplySpecified: providedSupplySpecified,
        notes: document.getElementById('ventCalcZoneNotes') ? String(document.getElementById('ventCalcZoneNotes').value || '').trim() : '',
        pz: pz,
        occupancy: occupancy
      };
    }

    function updateVentilationEntryReadouts() {
      function setFieldMuted(inputId, muted) {
        var inputNode = document.getElementById(inputId);
        if (!inputNode) return;
        var field = inputNode.closest('.ventilation-field');
        if (field) field.classList.toggle('is-muted', !!muted);
      }

      var draft = getVentilationEntryDraft();
      var occupancy = draft.occupancy;
      if (document.getElementById('ventCalcDefaultDensity')) {
        document.getElementById('ventCalcDefaultDensity').value = occupancy ? formatNumber(parseNonNegativeNumber(occupancy.default_occupant_density, 0), 2) : '';
      }
      if (document.getElementById('ventCalcPopulationDisplay')) {
        document.getElementById('ventCalcPopulationDisplay').value = occupancy ? formatNumber(parseNonNegativeNumber(draft.pz, 0), 2) : '0.00';
      }
      if (document.getElementById('ventCalcManualPopulation')) {
        document.getElementById('ventCalcManualPopulation').disabled = draft.populationMode !== 'MANUAL';
      }
      setFieldMuted('ventCalcManualPopulation', draft.populationMode !== 'MANUAL');
      setFieldMuted('ventCalcDefaultDensity', draft.populationMode === 'MANUAL');
      setFieldMuted('ventCalcPopulationDisplay', draft.populationMode === 'MANUAL');
      if (document.getElementById('ventCalcRpValue')) {
        document.getElementById('ventCalcRpValue').innerHTML = occupancy ? (formatNumber(parseNonNegativeNumber(occupancy.Rp, 0), 2) + ' CFM/person') : '—';
      }
      if (document.getElementById('ventCalcRaValue')) {
        document.getElementById('ventCalcRaValue').innerHTML = occupancy ? (formatNumber(parseNonNegativeNumber(occupancy.Ra, 0), 2) + ' CFM/ft²') : '—';
      }
      if (document.getElementById('ventCalcDensityValue')) {
        document.getElementById('ventCalcDensityValue').innerHTML = occupancy ? (formatNumber(parseNonNegativeNumber(occupancy.default_occupant_density, 0), 2) + ' people/1000 ft²') : '—';
      }
      if (document.getElementById('ventCalcAirClassValue')) {
        document.getElementById('ventCalcAirClassValue').innerHTML = occupancy ? ('Air Class ' + occupancy.air_class) : '—';
      }
      var rp = parseNonNegativeNumber(occupancy ? occupancy.Rp : 0, 0);
      var ra = parseNonNegativeNumber(occupancy ? occupancy.Ra : 0, 0);
      var vbz = (rp * draft.pz) + (ra * draft.area);
      var voz = draft.ez > 0 ? (vbz / draft.ez) : 0;
      var previewStatus = 'Not provided';
      if (draft.providedSupplySpecified) previewStatus = draft.providedSupply >= voz ? 'OK' : 'Under';
      if (document.getElementById('ventCalcPreviewPz')) document.getElementById('ventCalcPreviewPz').innerHTML = occupancy ? formatNumber(parseNonNegativeNumber(draft.pz, 0), 2) : '—';
      if (document.getElementById('ventCalcPreviewVbz')) document.getElementById('ventCalcPreviewVbz').innerHTML = occupancy ? (formatNumber(parseNonNegativeNumber(vbz, 0), 2) + ' CFM') : '—';
      if (document.getElementById('ventCalcPreviewEz')) document.getElementById('ventCalcPreviewEz').innerHTML = formatNumber(parseNonNegativeNumber(draft.ez, 0.8), 2);
      if (document.getElementById('ventCalcPreviewVoz')) document.getElementById('ventCalcPreviewVoz').innerHTML = occupancy ? (formatNumber(parseNonNegativeNumber(voz, 0), 2) + ' CFM') : '—';
      if (document.getElementById('ventCalcPreviewSupply')) {
        document.getElementById('ventCalcPreviewSupply').innerHTML = draft.providedSupplySpecified ? (formatNumber(parseNonNegativeNumber(draft.providedSupply, 0), 2) + ' CFM') : 'Not provided';
      }
      if (document.getElementById('ventCalcPreviewStatus')) {
        var statusNode = document.getElementById('ventCalcPreviewStatus');
        statusNode.innerHTML = previewStatus;
        statusNode.className = 'v ventilation-status-pill ' + (previewStatus === 'OK' ? 'status-ok' : (previewStatus === 'Under' ? 'status-under' : 'status-none'));
      }
    }

    function onVentilationEntryCategoryChanged() {
      updateVentilationEntryReadouts();
      calculateVentilation(false);
    }

    function onVentilationPopulationModeChanged() {
      updateVentilationEntryReadouts();
      calculateVentilation(false);
    }

    function onVentilationEzPresetChanged() {
      var preset = document.getElementById('ventCalcEzPreset') ? document.getElementById('ventCalcEzPreset').value : 'DEFAULT_DESIGN_BASIS_08';
      var ezNode = document.getElementById('ventCalcEz');
      if (!ezNode) return;
      for (var i = 0; i < ventilationEzCatalog.length; i++) {
        if (ventilationEzCatalog[i].key === preset && preset !== 'CUSTOM') {
          ezNode.value = formatNumber(parseNonNegativeNumber(ventilationEzCatalog[i].ez, 1), 2);
          break;
        }
      }
      updateVentilationEntryReadouts();
      calculateVentilation(false);
    }

    function updateVentilationEditUI() {
      var inEdit = !!ventilationEditingZoneId;
      var addBtn = document.getElementById('ventCalcAddUpdateBtn');
      var cancelBtn = document.getElementById('ventCalcCancelEditBtn');
      var editLabel = document.getElementById('ventCalcEditingLabel');
      var resetBtn = document.querySelector('#ventilationScreen .vent-action-reset');
      var dupBtn = document.querySelector('#ventilationScreen .vent-action-duplicate');
      if (addBtn) addBtn.value = inEdit ? 'UPDATE ZONE' : '+ ADD ZONE';
      if (cancelBtn) cancelBtn.style.display = inEdit ? '' : 'none';
      if (resetBtn) resetBtn.style.display = inEdit ? 'none' : '';
      if (dupBtn) dupBtn.style.display = inEdit ? 'none' : '';
      if (editLabel) {
        if (inEdit) {
          var zone = null;
          for (var i = 0; i < ventilationZones.length; i++) {
            if (ventilationZones[i].id === ventilationEditingZoneId) {
              zone = ventilationZones[i];
              break;
            }
          }
          editLabel.innerHTML = 'Editing zone: ' + (zone && zone.zoneName ? zone.zoneName : 'Zone');
          editLabel.style.display = '';
        } else {
          editLabel.style.display = 'none';
          editLabel.innerHTML = '';
        }
      }
    }

    function getNextVentilationZoneName() {
      ensureVentilationInitialized();
      return 'Zone ' + (ventilationZones.length + 1);
    }

    function resetVentilationInputForm(useSuggestedName) {
      ventilationEditingZoneId = '';
      updateVentilationEditUI();
      renderVentilationOccupancyOptions('');
      if (document.getElementById('ventCalcZoneName')) document.getElementById('ventCalcZoneName').value = useSuggestedName ? getNextVentilationZoneName() : '';
      if (document.getElementById('ventCalcSpaceName')) document.getElementById('ventCalcSpaceName').value = '';
      if (document.getElementById('ventCalcOccupancySearch')) document.getElementById('ventCalcOccupancySearch').value = '';
      if (document.getElementById('ventCalcOccupancy')) document.getElementById('ventCalcOccupancy').value = '';
      if (document.getElementById('ventCalcArea')) document.getElementById('ventCalcArea').value = 0;
      if (document.getElementById('ventCalcPopulationMode')) document.getElementById('ventCalcPopulationMode').value = 'DEFAULT';
      if (document.getElementById('ventCalcManualPopulation')) document.getElementById('ventCalcManualPopulation').value = 0;
      if (document.getElementById('ventCalcDefaultDensity')) document.getElementById('ventCalcDefaultDensity').value = '';
      if (document.getElementById('ventCalcPopulationDisplay')) document.getElementById('ventCalcPopulationDisplay').value = 0;
      if (document.getElementById('ventCalcEzPreset')) document.getElementById('ventCalcEzPreset').value = 'DEFAULT_DESIGN_BASIS_08';
      if (document.getElementById('ventCalcEz')) document.getElementById('ventCalcEz').value = 0.8;
      if (document.getElementById('ventCalcProvidedSupply')) document.getElementById('ventCalcProvidedSupply').value = '';
      if (document.getElementById('ventCalcZoneNotes')) document.getElementById('ventCalcZoneNotes').value = '';
      if (document.getElementById('ventCalcValidation')) document.getElementById('ventCalcValidation').innerHTML = '';
      updateVentilationEntryReadouts();
    }

    function cancelVentilationEdit() {
      resetVentilationInputForm(true);
    }

    function bindVentilationEntryLiveEvents() {
      var liveIds = ['ventCalcZoneName', 'ventCalcSpaceName', 'ventCalcArea', 'ventCalcManualPopulation', 'ventCalcEz', 'ventCalcProvidedSupply', 'ventCalcSystemPopulation', 'ventCalcZoneNotes'];
      for (var i = 0; i < liveIds.length; i++) {
        var node = document.getElementById(liveIds[i]);
        if (!node) continue;
        node.oninput = function () {
          updateVentilationEntryReadouts();
          calculateVentilation(false);
        };
        node.onchange = node.oninput;
      }
    }

    function addVentilationZoneFromInputs() {
      var draft = getVentilationEntryDraft();
      var zone = null;
      ensureVentilationInitialized();
      if (!draft.zoneName) {
        if (document.getElementById('ventCalcValidation')) document.getElementById('ventCalcValidation').innerHTML = 'Zone name is required.';
        return;
      }
      if (!draft.occupancyKey) {
        if (document.getElementById('ventCalcValidation')) document.getElementById('ventCalcValidation').innerHTML = 'Occupancy category is required.';
        return;
      }
      if (draft.area < 0) {
        if (document.getElementById('ventCalcValidation')) document.getElementById('ventCalcValidation').innerHTML = 'Area must be greater than or equal to 0.';
        return;
      }
      if (draft.ez <= 0) {
        if (document.getElementById('ventCalcValidation')) document.getElementById('ventCalcValidation').innerHTML = 'Ez must be greater than 0.';
        return;
      }
      if (ventilationEditingZoneId) {
        for (var i = 0; i < ventilationZones.length; i++) {
          if (ventilationZones[i].id === ventilationEditingZoneId) {
            zone = ventilationZones[i];
            break;
          }
        }
      }
      if (!zone) {
        zone = createDefaultVentilationZone(ventilationZones.length + 1);
        ventilationZones.push(zone);
      }
      zone.zoneName = draft.zoneName;
      zone.spaceName = draft.spaceName;
      zone.occupancyKey = draft.occupancyKey;
      zone.area = draft.area;
      zone.populationMode = draft.populationMode;
      zone.manualPopulation = draft.manualPopulation;
      zone.ezPreset = draft.ezPreset;
      zone.ez = draft.ez;
      zone.providedSupply = draft.providedSupply;
      zone.providedSupplySpecified = !!draft.providedSupplySpecified;
      zone.notes = draft.notes;
      ventilationEditingZoneId = '';
      updateVentilationEditUI();
      renderVentilationZonesTable();
      calculateVentilation(true);
      resetVentilationInputForm(true);
      calculateVentilation(false);
    }

    function duplicateVentilationZone(zoneId) {
      var source = null;
      var copy;
      for (var i = 0; i < ventilationZones.length; i++) {
        if (ventilationZones[i].id === zoneId) {
          source = ventilationZones[i];
          break;
        }
      }
      if (!source) return;
      copy = JSON.parse(JSON.stringify(source));
      copy.id = makeVentilationZoneId();
      copy.zoneName = (source.zoneName || 'Zone') + ' Copy';
      ventilationZones.push(copy);
      renderVentilationZonesTable();
      calculateVentilation(false);
    }

    function duplicateVentilationZoneFromEditor() {
      if (ventilationEditingZoneId) {
        duplicateVentilationZone(ventilationEditingZoneId);
        return;
      }
      if (ventilationZones && ventilationZones.length) duplicateVentilationZone(ventilationZones[0].id);
    }

    function editVentilationZone(zoneId) {
      var zone = null;
      for (var i = 0; i < ventilationZones.length; i++) {
        if (ventilationZones[i].id === zoneId) {
          zone = ventilationZones[i];
          break;
        }
      }
      if (!zone) return;
      ventilationEditingZoneId = zoneId;
      if (document.getElementById('ventCalcZoneName')) document.getElementById('ventCalcZoneName').value = zone.zoneName || '';
      if (document.getElementById('ventCalcSpaceName')) document.getElementById('ventCalcSpaceName').value = zone.spaceName || '';
      setVentilationOccupancySelection(zone.occupancyKey || '', '');
      if (document.getElementById('ventCalcArea')) document.getElementById('ventCalcArea').value = parseNonNegativeNumber(zone.area, 0);
      if (document.getElementById('ventCalcPopulationMode')) document.getElementById('ventCalcPopulationMode').value = zone.populationMode === 'MANUAL' ? 'MANUAL' : 'DEFAULT';
      if (document.getElementById('ventCalcManualPopulation')) document.getElementById('ventCalcManualPopulation').value = parseNonNegativeNumber(zone.manualPopulation, 0);
      if (document.getElementById('ventCalcEzPreset')) document.getElementById('ventCalcEzPreset').value = zone.ezPreset || 'DEFAULT_DESIGN_BASIS_08';
      if (document.getElementById('ventCalcEz')) document.getElementById('ventCalcEz').value = parseNonNegativeNumber(zone.ez, 0.8);
      if (document.getElementById('ventCalcProvidedSupply')) {
        document.getElementById('ventCalcProvidedSupply').value = zone.providedSupplySpecified ? formatNumber(parseNonNegativeNumber(zone.providedSupply, 0), 2) : '';
      }
      if (document.getElementById('ventCalcZoneNotes')) document.getElementById('ventCalcZoneNotes').value = zone.notes || '';
      updateVentilationEditUI();
      updateVentilationEntryReadouts();
    }

    function deleteVentilationZone(zoneId) {
      ventilationZones = ventilationZones.filter(function (z) { return z.id !== zoneId; });
      if (ventilationEditingZoneId === zoneId) ventilationEditingZoneId = '';
      updateVentilationEditUI();
      renderVentilationZonesTable();
      calculateVentilation(false);
    }

    function renderVentilationRunningTotals(summary) {
      var node = document.getElementById('ventilationRunningTotals');
      if (!node) return;
      if (!summary) {
        node.innerHTML = '';
        return;
      }
      node.innerHTML =
        '<div class="chip"><span class="k">Total Zones</span><span class="v">' + summary.zoneCount + '</span></div>' +
        '<div class="chip"><span class="k">Total Area</span><span class="v">' + formatNumber(summary.totalArea, 2) + ' ft²</span></div>' +
        '<div class="chip"><span class="k">Total Population</span><span class="v">' + formatNumber(summary.sumPz, 2) + '</span></div>' +
        '<div class="chip"><span class="k">Σ(Rp × Pz)</span><span class="v">' + formatNumber(summary.sumRpPz, 2) + '</span></div>' +
        '<div class="chip"><span class="k">Σ(Ra × Az)</span><span class="v">' + formatNumber(summary.sumRaAz, 2) + '</span></div>' +
        '<div class="chip"><span class="k">ΣVoz</span><span class="v">' + formatNumber(summary.sumVoz, 2) + ' CFM</span></div>';
    }

    function renderVentilationZonesTable() {
      var body = document.getElementById('ventilationZonesBody');
      var html = '';
      var calculated;
      var rows;
      if (!body) return;
      ensureVentilationInitialized();
      calculated = computeVentilationResult(false);
      rows = calculated && calculated.zones ? calculated.zones : [];
      if (!rows.length) {
        body.innerHTML = '<tr><td colspan="16" class="ventilation-empty-row">No zones added yet. Enter zone information above and click + ADD ZONE.</td></tr>';
        renderVentilationRunningTotals(calculated ? calculated.summary : null);
        return;
      }
      for (var i = 0; i < rows.length; i++) {
        var row = rows[i];
        html += '<tr>';
        html += '<td>' + (row.zoneName || ('Zone ' + (i + 1))) + '</td>';
        html += '<td>' + (row.spaceName || '') + '</td>';
        html += '<td>' + (row.occupancyCategory || '') + '</td>';
        html += '<td class="num">' + formatNumber(parseNonNegativeNumber(row.area, 0), 2) + '</td>';
        html += '<td>' + (row.populationMode === 'MANUAL'
          ? (formatNumber(parseNonNegativeNumber(row.pz, 0), 2) + ' people')
          : (formatNumber(parseNonNegativeNumber(row.density, 0), 2) + ' /1000 ft²')) + '</td>';
        html += '<td class="num">' + formatNumber(parseNonNegativeNumber(row.pz, 0), 2) + '</td>';
        html += '<td class="num">' + formatNumber(parseNonNegativeNumber(row.rp, 0), 2) + '</td>';
        html += '<td class="num">' + formatNumber(parseNonNegativeNumber(row.ra, 0), 2) + '</td>';
        html += '<td class="num">' + formatNumber(parseNonNegativeNumber(row.density, 0), 2) + '</td>';
        html += '<td class="num">' + formatNumber(parseNonNegativeNumber(row.ez, 0), 2) + '</td>';
        html += '<td class="num">' + formatNumber(parseNonNegativeNumber(row.vbz, 0), 2) + '</td>';
        html += '<td class="num">' + formatNumber(parseNonNegativeNumber(row.voz, 0), 2) + '</td>';
        html += '<td class="num">' + (row.providedSupplySpecified ? formatNumber(parseNonNegativeNumber(row.providedSupply, 0), 2) : '—') + '</td>';
        html += '<td class="' + (row.supplyStatus === 'OK' ? 'status-ok' : (row.supplyStatus === 'Under' ? 'status-under' : 'status-none')) + '">' + row.supplyStatus + '</td>';
        html += '<td>' + (row.notes || '') + '</td>';
        html += '<td><input type="button" class="btn cond-secondary-btn vent-row-action" value="Edit" onclick="editVentilationZone(\'' + row.id + '\');" /> <input type="button" class="btn cond-secondary-btn vent-row-action" value="Dup" onclick="duplicateVentilationZone(\'' + row.id + '\');" /> <input type="button" class="btn btn-small zone-danger-btn vent-row-action" value="Delete" onclick="deleteVentilationZone(\'' + row.id + '\');" /></td>';
        html += '</tr>';
      }
      body.innerHTML = html;
      renderVentilationRunningTotals(calculated ? calculated.summary : null);
    }

    function getVentilationSelectionState() {
      ensureVentilationInitialized();
      return {
        projectName: (hasActiveProject() && activeProject) ? String(activeProject.projectName || '') : '',
        systemType: document.getElementById('ventCalcSystemType') ? String(document.getElementById('ventCalcSystemType').value || 'SINGLE_ZONE') : 'SINGLE_ZONE',
        systemPopulationPs: parseNonNegativeNumber(document.getElementById('ventCalcSystemPopulation') ? document.getElementById('ventCalcSystemPopulation').value : 0, 0),
        occupancySearch: document.getElementById('ventCalcOccupancySearch') ? String(document.getElementById('ventCalcOccupancySearch').value || '') : '',
        zones: ventilationZones.map(function (z) {
          return {
            id: z.id,
            zoneName: z.zoneName || '',
            spaceName: z.spaceName || '',
            occupancyKey: z.occupancyKey || '',
            occupancyCategory: (getVentilationOccupancyByKey(z.occupancyKey) || {}).occupancy_category || '',
            area: parseNonNegativeNumber(z.area, 0),
            populationMode: z.populationMode === 'MANUAL' ? 'MANUAL' : 'DEFAULT',
            manualPopulation: parseNonNegativeNumber(z.manualPopulation, 0),
            ezPreset: z.ezPreset || 'DEFAULT_DESIGN_BASIS_08',
            ez: parseNonNegativeNumber(z.ez, 0.8),
            providedSupply: parseNonNegativeNumber(z.providedSupply, 0),
            providedSupplySpecified: !!z.providedSupplySpecified,
            notes: z.notes || ''
          };
        })
      };
    }

    function applyVentilationSelectionState(selection) {
      var zones = (selection && selection.zones && selection.zones.length) ? selection.zones : [];
      renderVentilationOccupancyOptions(selection.occupancySearch || '');
      renderVentilationEzOptions();
      if (document.getElementById('ventCalcSystemType')) document.getElementById('ventCalcSystemType').value = selection.systemType || 'SINGLE_ZONE';
      if (document.getElementById('ventCalcSystemPopulation')) document.getElementById('ventCalcSystemPopulation').value = parseNonNegativeNumber(selection.systemPopulationPs, 0);
      if (document.getElementById('ventCalcOccupancySearch')) document.getElementById('ventCalcOccupancySearch').value = selection.occupancySearch || '';
      ventilationZones = [];
      ventilationZoneIdSeq = 0;
      if (zones.length) {
        for (var i = 0; i < zones.length; i++) {
          var row = zones[i] || {};
          var zone = createDefaultVentilationZone(ventilationZones.length + 1);
          zone.id = row.id || makeVentilationZoneId();
          zone.zoneName = String(row.zoneName || ('Zone ' + (i + 1)));
          zone.spaceName = String(row.spaceName || '');
          zone.occupancyKey = resolveVentilationOccupancyKey(row.occupancyKey, row.occupancyCategory || row.occupancy_category);
          zone.area = parseNonNegativeNumber(row.area, 0);
          zone.populationMode = row.populationMode === 'MANUAL' ? 'MANUAL' : 'DEFAULT';
          zone.manualPopulation = parseNonNegativeNumber(row.manualPopulation, 0);
          zone.ezPreset = String(row.ezPreset || 'DEFAULT_DESIGN_BASIS_08');
          zone.ez = parseNonNegativeNumber(row.ez, 0.8);
          zone.providedSupply = parseNonNegativeNumber(row.providedSupply, 0);
          zone.providedSupplySpecified = !!row.providedSupplySpecified || parseNonNegativeNumber(row.providedSupply, 0) > 0;
          zone.notes = String(row.notes || '');
          ventilationZones.push(zone);
        }
      }
      ventilationEditingZoneId = '';
      updateVentilationEditUI();
      if (ventilationZones.length) {
        if (document.getElementById('ventCalcZoneName')) document.getElementById('ventCalcZoneName').value = ventilationZones[0].zoneName || '';
        if (document.getElementById('ventCalcSpaceName')) document.getElementById('ventCalcSpaceName').value = ventilationZones[0].spaceName || '';
        setVentilationOccupancySelection(ventilationZones[0].occupancyKey || '', selection.occupancySearch || '');
        if (document.getElementById('ventCalcArea')) document.getElementById('ventCalcArea').value = parseNonNegativeNumber(ventilationZones[0].area, 0);
        if (document.getElementById('ventCalcPopulationMode')) document.getElementById('ventCalcPopulationMode').value = ventilationZones[0].populationMode || 'DEFAULT';
        if (document.getElementById('ventCalcManualPopulation')) document.getElementById('ventCalcManualPopulation').value = parseNonNegativeNumber(ventilationZones[0].manualPopulation, 0);
        if (document.getElementById('ventCalcEzPreset')) document.getElementById('ventCalcEzPreset').value = ventilationZones[0].ezPreset || 'DEFAULT_DESIGN_BASIS_08';
        if (document.getElementById('ventCalcEz')) document.getElementById('ventCalcEz').value = parseNonNegativeNumber(ventilationZones[0].ez, 0.8);
        if (document.getElementById('ventCalcProvidedSupply')) {
          document.getElementById('ventCalcProvidedSupply').value = ventilationZones[0].providedSupplySpecified ? formatNumber(parseNonNegativeNumber(ventilationZones[0].providedSupply, 0), 2) : '';
        }
        if (document.getElementById('ventCalcZoneNotes')) document.getElementById('ventCalcZoneNotes').value = ventilationZones[0].notes || '';
      } else {
        resetVentilationInputForm(true);
      }
      updateVentilationEntryReadouts();
    }

    function computeVentilationResult(includeValidation) {
      var systemType = document.getElementById('ventCalcSystemType') ? document.getElementById('ventCalcSystemType').value : 'SINGLE_ZONE';
      var psInput = parseNonNegativeNumber(document.getElementById('ventCalcSystemPopulation') ? document.getElementById('ventCalcSystemPopulation').value : 0, 0);
      var rows = [];
      var warnings = [];
      var sumPz = 0;
      var sumVoz = 0;
      var sumRpPz = 0;
      var sumRaAz = 0;
      var sumVbz = 0;
      var totalArea = 0;
      ensureVentilationInitialized();

      for (var i = 0; i < ventilationZones.length; i++) {
        var z = ventilationZones[i];
        var occ = getVentilationOccupancyByKey(z.occupancyKey);
        var area = parseNonNegativeNumber(z.area, 0);
        var rp = parseNonNegativeNumber(occ ? occ.Rp : 0, 0);
        var ra = parseNonNegativeNumber(occ ? occ.Ra : 0, 0);
        var density = parseNonNegativeNumber(occ ? occ.default_occupant_density : 0, 0);
        var pz = (z.populationMode === 'MANUAL') ? parseNonNegativeNumber(z.manualPopulation, 0) : ((area * density) / 1000);
        var ez = parseNonNegativeNumber(z.ez, 0);
        var vbz = (rp * pz) + (ra * area);
        var voz = (ez > 0) ? (vbz / ez) : 0;
        var vpzMin = 1.5 * voz;
        var providedSupplySpecified = !!z.providedSupplySpecified;
        var providedSupply = providedSupplySpecified ? parseNonNegativeNumber(z.providedSupply, 0) : 0;
        var supplyStatus = '—';
        if (providedSupplySpecified) supplyStatus = providedSupply >= voz ? 'OK' : 'Under';
        if (!occ) warnings.push('Zone ' + (i + 1) + ': occupancy category is not selected.');
        if (area < 0) warnings.push('Zone ' + (i + 1) + ': area must be >= 0.');
        if (pz < 0) warnings.push('Zone ' + (i + 1) + ': population must be >= 0.');
        if (ez <= 0) warnings.push('Zone ' + (i + 1) + ': Ez must be greater than zero.');
        sumPz += pz;
        sumVoz += voz;
        sumVbz += vbz;
        sumRpPz += (rp * pz);
        sumRaAz += (ra * area);
        totalArea += area;
        rows.push({
          id: z.id,
          zoneName: z.zoneName || ('Zone ' + (i + 1)),
          spaceName: z.spaceName || '',
          occupancyKey: z.occupancyKey || '',
          occupancyCategory: occ ? occ.occupancy_category : '',
          categoryGroup: occ ? occ.category_group : '',
          rp: rp,
          ra: ra,
          density: density,
          airClass: occ ? occ.air_class : '',
          area: area,
          populationMode: z.populationMode,
          pz: roundNumber(pz, 4),
          ezLabel: (function () {
            for (var ex = 0; ex < ventilationEzCatalog.length; ex++) if (ventilationEzCatalog[ex].key === z.ezPreset) return ventilationEzCatalog[ex].label;
            return 'Manual custom Ez';
          })(),
          ez: roundNumber(ez, 4),
          vbz: roundNumber(vbz, 4),
          voz: roundNumber(voz, 4),
          vpzMin: roundNumber(vpzMin, 4),
          providedSupply: roundNumber(providedSupply, 4),
          providedSupplySpecified: providedSupplySpecified,
          supplyStatus: supplyStatus,
          notes: z.notes || '',
          details: {
            pzFormula: (z.populationMode === 'MANUAL')
              ? ('Pz = Manual = ' + formatNumber(pz, 2))
              : ('Pz = Az × density / 1000 = ' + formatNumber(area, 2) + ' × ' + formatNumber(density, 2) + ' / 1000 = ' + formatNumber(pz, 2)),
            vbzFormula: 'Vbz = Rp × Pz + Ra × Az = ' + formatNumber(rp, 2) + ' × ' + formatNumber(pz, 2) + ' + ' + formatNumber(ra, 2) + ' × ' + formatNumber(area, 2) + ' = ' + formatNumber(vbz, 2),
            vozFormula: 'Voz = Vbz / Ez = ' + formatNumber(vbz, 2) + ' / ' + formatNumber(ez, 2) + ' = ' + formatNumber(voz, 2),
            vpzFormula: 'Vpz-min = 1.5 × Voz = 1.5 × ' + formatNumber(voz, 2) + ' = ' + formatNumber(vpzMin, 2)
          }
        });
      }

      var ps = (psInput > 0) ? psInput : sumPz;
      var D = (sumPz > 0) ? (ps / sumPz) : 0;
      var Ev = (D < 0.6) ? (0.88 * D + 0.22) : 0.75;
      var Vou = (D * sumRpPz) + sumRaAz;
      var Vot = 0;
      if (systemType === 'SINGLE_ZONE') Vot = rows.length ? rows[0].voz : 0;
      else if (systemType === 'ALL_OA') Vot = sumVoz;
      else Vot = (Ev > 0 ? (Vou / Ev) : 0);

      if (systemType === 'MULTI_RECIRC' && ps > sumPz && sumPz > 0) warnings.push('System population Ps is greater than summed zone populations.');
      if (rows.length && systemType === 'MULTI_RECIRC' && sumPz <= 0) warnings.push('Sum of zone populations is zero. Diversity/system values are informational only.');
      if (psInput < 0) warnings.push('System population Ps must be >= 0.');

      var status = warnings.length ? 'incomplete' : 'ok';
      var message = warnings.length ? warnings[0] : 'Calculation complete.';
      if (!rows.length) {
        status = 'incomplete';
        message = 'No saved zones added yet.';
      }

      return {
        section: 'ventilation',
        status: status,
        message: message,
        systemType: systemType,
        zones: rows,
        summary: {
          zoneCount: rows.length,
          totalArea: roundNumber(totalArea, 4),
          sumPz: roundNumber(sumPz, 4),
          sumRpPz: roundNumber(sumRpPz, 4),
          sumRaAz: roundNumber(sumRaAz, 4),
          sumVbz: roundNumber(sumVbz, 4),
          sumVoz: roundNumber(sumVoz, 4)
        },
        ps: roundNumber(ps, 4),
        D: roundNumber(D, 4),
        Ev: roundNumber(Ev, 4),
        Vou: roundNumber(Vou, 4),
        Vot: roundNumber(Vot, 4),
        formulas: {
          diversity: 'D = Ps / ΣPz = ' + formatNumber(ps, 2) + ' / ' + formatNumber(sumPz, 2) + ' = ' + formatNumber(D, 3),
          vou: 'Vou = D × Σ(Rp × Pz) + Σ(Ra × Az) = ' + formatNumber(D, 3) + ' × ' + formatNumber(sumRpPz, 2) + ' + ' + formatNumber(sumRaAz, 2) + ' = ' + formatNumber(Vou, 2),
          ev: (D < 0.6 ? ('Ev = 0.88 × D + 0.22 = 0.88 × ' + formatNumber(D, 3) + ' + 0.22 = ' + formatNumber(Ev, 3)) : 'Ev = 0.75 (D >= 0.60)'),
          vot: 'Vot = Vou / Ev = ' + formatNumber(Vou, 2) + ' / ' + formatNumber(Ev, 3) + ' = ' + formatNumber(Vot, 2)
        },
        warnings: warnings
      };
    }

    function renderVentilationResultHtml(result) {
      if (!result) return '<div class="result-placeholder">Select occupancy and enter zone data to calculate required outdoor air.</div>';
      if (!result.summary) {
        result.summary = { zoneCount: 0, sumVbz: 0, sumVoz: 0 };
      }
      var hasZones = !!(result.summary && result.summary.zoneCount);
      var statusText = hasZones ? ((result.status === 'ok') ? 'Calculated' : 'Incomplete') : 'No saved zones';
      var systemLabel = (result.systemType === 'MULTI_RECIRC') ? 'Multiple-zone recirculating system' : ((result.systemType === 'ALL_OA') ? '100% outdoor air system' : 'Single-zone system');
      var html = '';
      html += '<div class="ventilation-pill-grid">';
      html += '<div class="ventilation-pill"><span class="k">System Type</span><span class="v">' + systemLabel + '</span></div>';
      html += '<div class="ventilation-pill"><span class="k">Final Intake Vot</span><span class="v">' + formatNumber(result.Vot, 2) + ' CFM</span></div>';
      html += '<div class="ventilation-pill"><span class="k">Final Intake Vot (L/s)</span><span class="v">' + formatNumber(cfmToLs(result.Vot), 2) + ' L/s</span></div>';
      html += '<div class="ventilation-pill"><span class="k">Status</span><span class="v">' + statusText + '</span></div>';
      html += '</div>';
      html += '<table class="ventilation-summary-table"><tbody>';
      html += '<tr><th>ΣVbz</th><td class="num">' + formatNumber(result.summary.sumVbz, 2) + ' CFM</td><th>ΣVoz</th><td class="num">' + formatNumber(result.summary.sumVoz, 2) + ' CFM</td></tr>';
      html += '</tbody></table>';
      if (!hasZones) html += '<div class="subtle ventilation-result-note">Results summarize saved zones only. Use Current Zone Preview for unsaved input.</div>';
      if (result.warnings && result.warnings.length) html += '<div class="warning-line">' + result.warnings.join(' | ') + '</div>';
      return html;
    }

    function renderVentilationCalculationDetails(result) {
      var node = document.getElementById('ventilationCalculationDetails');
      var html = '';
      if (!node) return;
      if (!result || !result.zones || !result.zones.length) {
        node.innerHTML = 'Zone and system formula details will appear after calculation.';
        return;
      }
      for (var i = 0; i < result.zones.length; i++) {
        var z = result.zones[i];
        html += '<div style="margin-bottom:7px;"><strong>' + (z.zoneName || ('Zone ' + (i + 1))) + '</strong><br>' + z.details.pzFormula + '<br>' + z.details.vbzFormula + '<br>' + z.details.vozFormula + '<br>' + z.details.vpzFormula;
        if (z.providedSupplySpecified) html += '<br>Provided Supply = ' + formatNumber(parseNonNegativeNumber(z.providedSupply, 0), 2) + ' CFM → ' + (z.supplyStatus || '—');
        html += '</div>';
      }
      html += '<div><strong>System</strong><br>' + result.formulas.diversity + '<br>' + result.formulas.vou + '<br>' + result.formulas.ev + '<br>' + result.formulas.vot + '</div>';
      node.innerHTML = html;
    }

    function calculateVentilation(commitResults) {
      var validationNode = document.getElementById('ventCalcValidation');
      var resultsNode = document.getElementById('ventilationResults');
      var result = computeVentilationResult(true);
      if (validationNode) validationNode.innerHTML = (result.warnings && result.warnings.length) ? result.warnings[0] : '';
      if (resultsNode) resultsNode.innerHTML = renderVentilationResultHtml(result);
      renderVentilationCalculationDetails(result);
      latestVentilationSnapshot = {
        section: 'ventilation',
        selection: getVentilationSelectionState(),
        result: result
      };
      renderVentilationZonesTable();
    }

    function resetVentilationSection() {
      latestVentilationSnapshot = null;
      renderVentilationEzOptions();
      ensureVentilationInitialized();
      resetVentilationInputForm(true);
      bindVentilationEntryLiveEvents();
      renderVentilationZonesTable();
      calculateVentilation(false);
    }

    function clearVentilationSavedZones() {
      latestVentilationSnapshot = null;
      ventilationEditingZoneId = '';
      ventilationZoneIdSeq = 0;
      ventilationZones = [];
      renderVentilationEzOptions();
      resetVentilationInputForm(true);
      bindVentilationEntryLiveEvents();
      renderVentilationZonesTable();
      calculateVentilation(false);
    }

    async function syncAppMetaFromBackend(attempt) {
      var tryCount = parseInt(attempt, 10);
      if (isNaN(tryCount) || tryCount < 0) tryCount = 0;
      var versionCandidate = '';

      try {
        if (window.pywebview && window.pywebview.api && window.pywebview.api.get_app_meta) {
          var meta = await window.pywebview.api.get_app_meta();
          if (meta && meta.ok) {
            if (!versionCandidate && typeof meta.app_version === 'string' && meta.app_version.trim()) {
              versionCandidate = meta.app_version.trim();
            }
            if (typeof meta.project_schema_version === 'number' && meta.project_schema_version > 0) APP_SCHEMA_VERSION = meta.project_schema_version;
          }
        }
      } catch (e) {
        // Keep UI stable; hide version text if unavailable instead of showing stale values.
      }

      if (versionCandidate) APP_VERSION = versionCandidate;

      if (!APP_VERSION && tryCount < 12) {
        window.setTimeout(function () { syncAppMetaFromBackend(tryCount + 1); }, 300);
      }

      renderCurrentProjectCard();
    }
    function initPage() {
      activeProject = null;
      savedProjects = [];
      resetForm();
      renderCurrentProjectCard();
      syncAppMetaFromBackend(0);
      window.setTimeout(function () { syncAppMetaFromBackend(0); }, 700);
      seedLegacyGasTables();
      initGasMaterialOptions();
      initSolarCatalog();
      if (document.getElementById('gasCodeBasis')) document.getElementById('gasCodeBasis').value = 'IPC';
      if (document.getElementById('gasGasType')) document.getElementById('gasGasType').value = 'NATURAL_GAS';
      if (document.getElementById('gasMaterial')) document.getElementById('gasMaterial').value = 'SCHEDULE_40_METALLIC';
      refreshGasCriteriaDropdowns();
      if (document.getElementById('gasDemandMode')) document.getElementById('gasDemandMode').value = 'CFH';
      if (document.getElementById('gasHeatingValue')) document.getElementById('gasHeatingValue').value = 0;
      clearGasLineRows();
      addGasLineRow({ id: makeGasLineId(), label: 'Line 1', demandValue: 0, runLength: 0 });
      if (document.getElementById('gasResults')) document.getElementById('gasResults').innerHTML = 'Enter gas inputs and press Calculate.';
      resetWsfuSection();
      resetFixtureUnitSection();
      resetVentilationSection();
      resetDuctSection();
      resetDuctStaticSection();
      if (document.getElementById('solarBuildingSf')) document.getElementById('solarBuildingSf').value = 0;
      if (document.getElementById('solarZip')) document.getElementById('solarZip').value = '';
      if (document.getElementById('solarClimateZone')) document.getElementById('solarClimateZone').value = '';
      if (document.getElementById('solarUseSaraPath')) document.getElementById('solarUseSaraPath').checked = false;
      if (document.getElementById('solarSara')) document.getElementById('solarSara').value = '';
      if (document.getElementById('solarDValue')) document.getElementById('solarDValue').value = 0;
      if (document.getElementById('solarExempt')) document.getElementById('solarExempt').checked = false;
      if (document.getElementById('batteryExempt')) document.getElementById('batteryExempt').checked = false;
      if (document.getElementById('solarNotes')) document.getElementById('solarNotes').value = '';
      solarClimateManualOverride = false;
      if (solarZipLookupTimer) {
        window.clearTimeout(solarZipLookupTimer);
        solarZipLookupTimer = null;
      }
      if (solarAutoCalcTimer) {
        window.clearTimeout(solarAutoCalcTimer);
        solarAutoCalcTimer = null;
      }
      onSolarSaraToggle();
      if (document.getElementById('solarResults')) document.getElementById('solarResults').innerHTML = 'Enter worksheet inputs to see live Solar results (or press CALCULATE).';
      updateCondensatePreview();
      updateVentTotalsAndUI();
      calculateGasPipe(false);
      updateWsfuTotalsAndUI();
      calculateDuct(false);
      updateSolarPreview();
      latestWsfuSnapshot = null;
      latestDuctSnapshot = null;
      latestDuctStaticSnapshot = null;
      latestSolarSnapshot = null;
      latestRefrigerantSnapshot = null;
      latestVentilationSnapshot = null;
      setActiveModule('home');
      updateHomeDashboardFit();
    }
    window.addEventListener('pywebviewready', function () { syncAppMetaFromBackend(0); });
    window.addEventListener('resize', function () { updateHomeDashboardFit(); });


