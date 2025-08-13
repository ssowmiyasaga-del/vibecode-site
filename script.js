/* Utilities */
const $ = (selector, root = document) => root.querySelector(selector);
const $$ = (selector, root = document) => Array.from(root.querySelectorAll(selector));

/* Data: comprehensive Excel resources with detailed information */
const resources = [
  // Beginner Essentials
  { 
    title: 'Excel Interface Mastery', 
    type: 'Tutorial', 
    tags: ['Beginner'], 
    description: 'Master Excel\'s interface: ribbon navigation, quick access toolbar customization, and essential keyboard shortcuts. Learn to navigate efficiently and customize your workspace for maximum productivity.',
    difficulty: 'Beginner', 
    duration: '20 min',
    detailedInfo: {
      whatYouLearn: ['Ribbon navigation and customization', 'Quick Access Toolbar setup', 'Essential keyboard shortcuts', 'Worksheet navigation techniques', 'Customizing the Excel interface'],
      prerequisites: 'No prior Excel experience required',
      includes: ['Step-by-step video tutorials', 'Interactive exercises', 'Downloadable practice files', 'Keyboard shortcut cheat sheet'],
      outcomes: 'Navigate Excel confidently and customize your workspace for efficiency'
    }
  },
  { 
    title: 'Basic Formulas Starter Kit', 
    type: 'Tutorial', 
    tags: ['Beginner', 'Formulas'], 
    description: 'Master fundamental Excel formulas: SUM, AVERAGE, COUNT, MIN, MAX with real-world examples. Build confidence with basic calculations and understand formula syntax.',
    difficulty: 'Beginner', 
    duration: '25 min',
    detailedInfo: {
      whatYouLearn: ['SUM function for adding numbers', 'AVERAGE function for calculating means', 'COUNT functions for data analysis', 'MIN and MAX for finding extremes', 'Basic formula syntax and cell references'],
      prerequisites: 'Basic Excel interface knowledge',
      includes: ['Practice datasets', 'Formula examples', 'Common mistakes to avoid', 'Real-world scenarios'],
      outcomes: 'Confidently use basic Excel formulas for everyday calculations'
    }
  },
  { 
    title: 'Cell Formatting Tricks', 
    type: 'Tutorial', 
    tags: ['Beginner'], 
    description: 'Transform your spreadsheets with professional formatting: number formats, conditional formatting, and styling techniques that make data easy to read and understand.',
    difficulty: 'Beginner', 
    duration: '15 min',
    detailedInfo: {
      whatYouLearn: ['Number formatting (currency, percentages, dates)', 'Conditional formatting rules', 'Cell styles and themes', 'Professional table formatting', 'Custom number formats'],
      prerequisites: 'Basic Excel knowledge',
      includes: ['Formatting templates', 'Color scheme guidelines', 'Professional examples', 'Quick formatting shortcuts'],
      outcomes: 'Create professional-looking spreadsheets with clear, readable formatting'
    }
  },
  { 
    title: 'Data Entry Best Practices', 
    type: 'Tutorial', 
    tags: ['Beginner'], 
    description: 'Learn efficient data entry techniques, validation rules, and error prevention strategies. Build robust spreadsheets from the ground up.',
    difficulty: 'Beginner', 
    duration: '18 min',
    detailedInfo: {
      whatYouLearn: ['Efficient data entry techniques', 'Data validation rules', 'Error prevention strategies', 'Consistent data formatting', 'Data entry shortcuts'],
      prerequisites: 'Basic Excel interface knowledge',
      includes: ['Data validation templates', 'Error checking tools', 'Best practice guidelines', 'Common pitfalls to avoid'],
      outcomes: 'Enter data efficiently and prevent common errors in your spreadsheets'
    }
  },
  
  // Essential Formulas
  { 
    title: 'IF Function Mastery', 
    type: 'Tutorial', 
    tags: ['Formulas'], 
    description: 'Master conditional logic with IF, IFS, and SWITCH functions. Learn to create dynamic formulas that respond to different conditions and scenarios.',
    difficulty: 'Intermediate', 
    duration: '30 min',
    detailedInfo: {
      whatYouLearn: ['IF function syntax and logic', 'Nested IF statements', 'IFS function for multiple conditions', 'SWITCH function alternatives', 'Real-world business scenarios'],
      prerequisites: 'Basic Excel formulas knowledge',
      includes: ['Business case studies', 'Complex scenario examples', 'Performance optimization tips', 'Common IF function mistakes'],
      outcomes: 'Create dynamic, conditional formulas for complex business logic'
    }
  },
  { 
    title: 'VLOOKUP vs XLOOKUP Guide', 
    type: 'Tutorial', 
    tags: ['Formulas'], 
    description: 'Compare and master both VLOOKUP and XLOOKUP functions. Understand when to use each, their limitations, and advanced lookup techniques.',
    difficulty: 'Intermediate', 
    duration: '35 min',
    detailedInfo: {
      whatYouLearn: ['VLOOKUP function syntax and limitations', 'XLOOKUP modern alternative', 'When to use each function', 'Error handling techniques', 'Advanced lookup scenarios'],
      prerequisites: 'Basic Excel formulas knowledge',
      includes: ['Lookup table examples', 'Error handling strategies', 'Performance comparisons', 'Migration guide from VLOOKUP to XLOOKUP'],
      outcomes: 'Choose the right lookup function and handle complex data retrieval scenarios'
    }
  },
  { 
    title: 'Text Functions Complete Guide', 
    type: 'Tutorial', 
    tags: ['Formulas'], 
    description: 'Master text manipulation with LEFT, RIGHT, MID, CONCATENATE, and TEXT functions. Clean and transform text data efficiently.',
    difficulty: 'Intermediate', 
    duration: '28 min',
    detailedInfo: {
      whatYouLearn: ['LEFT, RIGHT, MID for text extraction', 'CONCATENATE and & operator', 'TEXT function for formatting', 'Text cleaning techniques', 'Combining text functions'],
      prerequisites: 'Basic Excel formulas knowledge',
      includes: ['Text cleaning templates', 'Common text scenarios', 'Performance optimization', 'Text function combinations'],
      outcomes: 'Efficiently manipulate and clean text data in Excel'
    }
  },
  { 
    title: 'Date & Time Calculations', 
    type: 'Tutorial', 
    tags: ['Formulas'], 
    description: 'Master date and time functions for calculating durations, working with time zones, and handling complex date scenarios in business applications.',
    difficulty: 'Intermediate', 
    duration: '32 min',
    detailedInfo: {
      whatYouLearn: ['TODAY and NOW functions', 'DATEDIF for date differences', 'EOMONTH and date calculations', 'Time zone conversions', 'Business date scenarios'],
      prerequisites: 'Basic Excel formulas knowledge',
      includes: ['Date calculation templates', 'Business calendar examples', 'Time tracking scenarios', 'Date validation techniques'],
      outcomes: 'Handle complex date and time calculations for business applications'
    }
  },
  
  // Advanced Formulas
  { 
    title: 'Dynamic Arrays Revolution', 
    type: 'Tutorial', 
    tags: ['Formulas'], 
    description: 'Explore the future of Excel formulas with FILTER, SORT, UNIQUE, and SEQUENCE functions. Transform how you work with data arrays.',
    difficulty: 'Advanced', 
    duration: '40 min',
    detailedInfo: {
      whatYouLearn: ['FILTER function for dynamic filtering', 'SORT function for dynamic sorting', 'UNIQUE function for removing duplicates', 'SEQUENCE function for generating sequences', 'Combining dynamic array functions'],
      prerequisites: 'Intermediate Excel formulas knowledge',
      includes: ['Dynamic array examples', 'Performance benefits', 'Compatibility considerations', 'Advanced array scenarios'],
      outcomes: 'Leverage dynamic arrays for powerful, flexible data analysis'
    }
  },
  { 
    title: 'Array Formulas Deep Dive', 
    type: 'Tutorial', 
    tags: ['Formulas'], 
    description: 'Master SUMPRODUCT, INDEX/MATCH, and complex array operations. Learn advanced techniques for sophisticated data analysis.',
    difficulty: 'Advanced', 
    duration: '45 min',
    detailedInfo: {
      whatYouLearn: ['SUMPRODUCT for weighted calculations', 'INDEX/MATCH combinations', 'Complex array operations', 'Performance optimization', 'Advanced lookup scenarios'],
      prerequisites: 'Intermediate Excel formulas knowledge',
      includes: ['Complex calculation examples', 'Performance optimization tips', 'Array formula best practices', 'Advanced business scenarios'],
      outcomes: 'Create sophisticated array formulas for complex data analysis'
    }
  },
  { 
    title: 'LET Function & Variables', 
    type: 'Tutorial', 
    tags: ['Formulas'], 
    description: 'Simplify complex formulas with the LET function and named variables. Make your formulas more readable and maintainable.',
    difficulty: 'Advanced', 
    duration: '25 min',
    detailedInfo: {
      whatYouLearn: ['LET function syntax', 'Named variables in formulas', 'Formula simplification techniques', 'Performance improvements', 'Complex formula breakdown'],
      prerequisites: 'Intermediate Excel formulas knowledge',
      includes: ['Complex formula examples', 'Performance comparisons', 'Best practices', 'Formula maintenance tips'],
      outcomes: 'Create cleaner, more maintainable complex formulas'
    }
  },
  
  // Data Analysis
  { 
    title: 'PivotTables from Scratch', 
    type: 'Tutorial', 
    tags: ['Analysis'], 
    description: 'Create powerful reports with PivotTables using drag-and-drop simplicity. Transform raw data into meaningful insights quickly.',
    difficulty: 'Intermediate', 
    duration: '35 min',
    detailedInfo: {
      whatYouLearn: ['PivotTable creation and setup', 'Field arrangement and formatting', 'Basic calculations and summaries', 'Filtering and sorting', 'PivotTable best practices'],
      prerequisites: 'Basic Excel knowledge',
      includes: ['Sample datasets', 'PivotTable templates', 'Common scenarios', 'Troubleshooting guide'],
      outcomes: 'Create effective PivotTables for data analysis and reporting'
    }
  },
  { 
    title: 'PivotTable Advanced Features', 
    type: 'Tutorial', 
    tags: ['Analysis'], 
    description: 'Master advanced PivotTable features: calculated fields, grouping, custom formatting, and slicers for interactive dashboards.',
    difficulty: 'Advanced', 
    duration: '40 min',
    detailedInfo: {
      whatYouLearn: ['Calculated fields and items', 'Grouping and ungrouping data', 'Custom formatting and styles', 'Slicers and timelines', 'Advanced filtering techniques'],
      prerequisites: 'Basic PivotTable knowledge',
      includes: ['Advanced PivotTable examples', 'Dashboard templates', 'Interactive features guide', 'Performance optimization'],
      outcomes: 'Create sophisticated PivotTable reports with advanced features'
    }
  },
  { 
    title: 'Power Query Basics', 
    type: 'Tutorial', 
    tags: ['Power Query', 'Analysis'], 
    description: 'Import, clean, and transform data without formulas using Power Query. Master data preparation for analysis.',
    difficulty: 'Intermediate', 
    duration: '30 min',
    detailedInfo: {
      whatYouLearn: ['Power Query interface', 'Data import techniques', 'Data cleaning and transformation', 'Merging and appending data', 'Query editor basics'],
      prerequisites: 'Basic Excel knowledge',
      includes: ['Sample data files', 'Transformation examples', 'Common data cleaning scenarios', 'Power Query templates'],
      outcomes: 'Efficiently prepare and clean data for analysis using Power Query'
    }
  },
  { 
    title: 'Data Validation Mastery', 
    type: 'Tutorial', 
    tags: ['Analysis'], 
    description: 'Create dropdown lists, custom rules, and error handling with data validation. Ensure data integrity and user-friendly input.',
    difficulty: 'Intermediate', 
    duration: '22 min',
    detailedInfo: {
      whatYouLearn: ['Dropdown list creation', 'Custom validation rules', 'Error message customization', 'Input restrictions', 'Validation best practices'],
      prerequisites: 'Basic Excel knowledge',
      includes: ['Validation templates', 'Custom rule examples', 'Error handling strategies', 'User experience guidelines'],
      outcomes: 'Create robust data validation systems for user input'
    }
  },
  
  // Templates & Practical Examples
  { 
    title: 'Personal Budget Template', 
    type: 'Template', 
    tags: ['Beginner'], 
    description: 'Complete personal budget tracker with income/expense categories, charts, and monthly summaries. Perfect for financial planning.',
    difficulty: 'Beginner', 
    duration: '45 min',
    detailedInfo: {
      whatYouLearn: ['Budget structure and categories', 'Income and expense tracking', 'Chart creation and formatting', 'Monthly summaries and analysis', 'Budget vs actual comparisons'],
      prerequisites: 'Basic Excel knowledge',
      includes: ['Complete budget template', 'Chart examples', 'Category suggestions', 'Monthly analysis tools'],
      outcomes: 'Create and maintain a comprehensive personal budget system'
    }
  },
  { 
    title: 'Sales Dashboard Template', 
    type: 'Template', 
    tags: ['Analysis'], 
    description: 'Professional sales dashboard with PivotTables, charts, and KPI tracking. Monitor sales performance and trends effectively.',
    difficulty: 'Intermediate', 
    duration: '60 min',
    detailedInfo: {
      whatYouLearn: ['Dashboard design principles', 'KPI tracking and visualization', 'PivotTable integration', 'Chart creation and formatting', 'Dynamic data updates'],
      prerequisites: 'Basic Excel and PivotTable knowledge',
      includes: ['Complete dashboard template', 'KPI calculation examples', 'Chart templates', 'Data update procedures'],
      outcomes: 'Create professional sales dashboards for business reporting'
    }
  },
  { 
    title: 'Project Tracker Template', 
    type: 'Template', 
    tags: ['Analysis'], 
    description: 'Gantt-style project management system with progress tracking, milestones, and resource allocation. Manage projects effectively.',
    difficulty: 'Intermediate', 
    duration: '50 min',
    detailedInfo: {
      whatYouLearn: ['Project timeline creation', 'Progress tracking methods', 'Milestone management', 'Resource allocation tracking', 'Project status reporting'],
      prerequisites: 'Basic Excel knowledge',
      includes: ['Project tracker template', 'Timeline examples', 'Progress calculation formulas', 'Status reporting tools'],
      outcomes: 'Create comprehensive project tracking systems'
    }
  },
  { 
    title: 'Inventory Management System', 
    type: 'Template', 
    tags: ['Analysis'], 
    description: 'Complete inventory tracking system with stock levels, reorder alerts, and detailed reporting. Perfect for small business inventory management.',
    difficulty: 'Advanced', 
    duration: '75 min',
    detailedInfo: {
      whatYouLearn: ['Inventory tracking structure', 'Stock level monitoring', 'Reorder point calculations', 'Inventory valuation methods', 'Detailed reporting systems'],
      prerequisites: 'Intermediate Excel knowledge',
      includes: ['Complete inventory system', 'Alert formulas', 'Valuation templates', 'Reporting dashboards'],
      outcomes: 'Create a comprehensive inventory management system'
    }
  },
  
  // VBA & Automation
  { 
    title: 'VBA Fundamentals', 
    type: 'Tutorial', 
    tags: ['VBA'], 
    description: 'Introduction to VBA programming: macros, procedures, and basic syntax. Start your automation journey with Excel VBA.',
    difficulty: 'Advanced', 
    duration: '50 min',
    detailedInfo: {
      whatYouLearn: ['VBA editor and interface', 'Basic VBA syntax', 'Sub procedures and functions', 'Variables and data types', 'Basic automation concepts'],
      prerequisites: 'Intermediate Excel knowledge',
      includes: ['VBA code examples', 'Practice exercises', 'Debugging techniques', 'Best practices guide'],
      outcomes: 'Write basic VBA code for Excel automation'
    }
  },
  { 
    title: 'Macro Recorder Guide', 
    type: 'Tutorial', 
    tags: ['VBA'], 
    description: 'Learn to record and edit macros for repetitive tasks. Transform manual processes into automated workflows.',
    difficulty: 'Advanced', 
    duration: '35 min',
    detailedInfo: {
      whatYouLearn: ['Macro recorder usage', 'Recording best practices', 'Editing recorded macros', 'Macro security settings', 'Common automation scenarios'],
      prerequisites: 'Basic Excel knowledge',
      includes: ['Macro examples', 'Recording guidelines', 'Security best practices', 'Troubleshooting tips'],
      outcomes: 'Create and edit macros for task automation'
    }
  },
  { 
    title: 'UserForms & Interfaces', 
    type: 'Tutorial', 
    tags: ['VBA'], 
    description: 'Create custom dialog boxes and user interfaces with VBA UserForms. Build professional applications within Excel.',
    difficulty: 'Advanced', 
    duration: '55 min',
    detailedInfo: {
      whatYouLearn: ['UserForm design and creation', 'Control properties and events', 'Form validation techniques', 'Data input and processing', 'Professional interface design'],
      prerequisites: 'Basic VBA knowledge',
      includes: ['UserForm examples', 'Control reference guide', 'Validation techniques', 'Design best practices'],
      outcomes: 'Create professional user interfaces with VBA UserForms'
    }
  },
  
  // Quick Reference Guides
  { 
    title: 'Formula Cheat Sheet', 
    type: 'Reference', 
    tags: ['Formulas'], 
    description: 'Quick reference guide for 50+ essential Excel formulas with syntax, examples, and usage tips. Perfect for daily reference.',
    difficulty: 'Beginner', 
    duration: '10 min',
    detailedInfo: {
      whatYouLearn: ['Formula syntax patterns', 'Common function categories', 'Quick reference techniques', 'Formula troubleshooting', 'Efficiency tips'],
      prerequisites: 'Basic Excel knowledge',
      includes: ['Complete formula reference', 'Syntax examples', 'Usage tips', 'Troubleshooting guide'],
      outcomes: 'Quick access to essential Excel formulas and syntax'
    }
  },
  { 
    title: 'Keyboard Shortcuts Guide', 
    type: 'Reference', 
    tags: ['Beginner'], 
    description: '100+ time-saving keyboard shortcuts for Excel navigation, editing, formatting, and formulas. Boost your productivity significantly.',
    difficulty: 'Beginner', 
    duration: '15 min',
    detailedInfo: {
      whatYouLearn: ['Navigation shortcuts', 'Editing and formatting shortcuts', 'Formula shortcuts', 'Custom shortcut creation', 'Productivity techniques'],
      prerequisites: 'Basic Excel knowledge',
      includes: ['Complete shortcut reference', 'Category organization', 'Practice exercises', 'Customization guide'],
      outcomes: 'Significantly improve Excel productivity with keyboard shortcuts'
    }
  },
  { 
    title: 'Error Handling Guide', 
    type: 'Reference', 
    tags: ['Formulas'], 
    description: 'Comprehensive guide to common Excel formula errors, their causes, and solutions. Troubleshoot formula problems effectively.',
    difficulty: 'Intermediate', 
    duration: '20 min',
    detailedInfo: {
      whatYouLearn: ['Common error types and meanings', 'Error prevention techniques', 'Debugging strategies', 'Formula troubleshooting', 'Best practices'],
      prerequisites: 'Basic Excel formulas knowledge',
      includes: ['Error reference guide', 'Troubleshooting examples', 'Prevention strategies', 'Debugging tools'],
      outcomes: 'Effectively troubleshoot and prevent Excel formula errors'
    }
  },
  { 
    title: 'Performance Optimization', 
    type: 'Reference', 
    tags: ['Analysis'], 
    description: 'Speed up slow workbooks and optimize Excel performance. Learn techniques for handling large datasets efficiently.',
    difficulty: 'Advanced', 
    duration: '25 min',
    detailedInfo: {
      whatYouLearn: ['Performance bottlenecks identification', 'Formula optimization techniques', 'Data structure improvements', 'Memory management', 'Calculation optimization'],
      prerequisites: 'Intermediate Excel knowledge',
      includes: ['Performance analysis tools', 'Optimization examples', 'Best practices guide', 'Troubleshooting techniques'],
      outcomes: 'Optimize Excel workbooks for maximum performance'
    }
  }
];



// Excel Formula Examples Database
const formulaExamples = {
  'Basic Math': [
    { name: 'SUM', syntax: '=SUM(A1:A10)', description: 'Adds all numbers in a range', example: '=SUM(B2:B15)' },
    { name: 'AVERAGE', syntax: '=AVERAGE(A1:A10)', description: 'Calculates the mean of numbers', example: '=AVERAGE(C2:C20)' },
    { name: 'COUNT', syntax: '=COUNT(A1:A10)', description: 'Counts cells with numbers', example: '=COUNT(D2:D25)' },
    { name: 'MAX/MIN', syntax: '=MAX(A1:A10)', description: 'Finds highest/lowest value', example: '=MAX(E2:E30)' }
  ],
  'Text Functions': [
    { name: 'LEFT', syntax: '=LEFT(text, num_chars)', description: 'Extracts characters from left', example: '=LEFT(A2, 3)' },
    { name: 'RIGHT', syntax: '=RIGHT(text, num_chars)', description: 'Extracts characters from right', example: '=RIGHT(A2, 2)' },
    { name: 'MID', syntax: '=MID(text, start_num, num_chars)', description: 'Extracts from middle', example: '=MID(A2, 2, 4)' },
    { name: 'CONCATENATE', syntax: '=CONCATENATE(text1, text2)', description: 'Joins text together', example: '=CONCATENATE(A2, " ", B2)' }
  ],
  'Lookup Functions': [
    { name: 'VLOOKUP', syntax: '=VLOOKUP(lookup_value, table_array, col_index, [range_lookup])', description: 'Vertical lookup in table', example: '=VLOOKUP(A2, B2:D10, 3, FALSE)' },
    { name: 'XLOOKUP', syntax: '=XLOOKUP(lookup_value, lookup_array, return_array)', description: 'Modern lookup function', example: '=XLOOKUP(A2, B2:B10, C2:C10)' },
    { name: 'INDEX/MATCH', syntax: '=INDEX(return_range, MATCH(lookup_value, lookup_range, 0))', description: 'Flexible lookup combination', example: '=INDEX(C2:C10, MATCH(A2, B2:B10, 0))' }
  ],
  'Logical Functions': [
    { name: 'IF', syntax: '=IF(logical_test, value_if_true, value_if_false)', description: 'Conditional logic', example: '=IF(A2>100, "High", "Low")' },
    { name: 'IFS', syntax: '=IFS(logical_test1, value1, logical_test2, value2)', description: 'Multiple conditions', example: '=IFS(A2>100, "High", A2>50, "Medium", TRUE, "Low")' },
    { name: 'AND/OR', syntax: '=AND(logical1, logical2)', description: 'Multiple logical tests', example: '=AND(A2>50, B2<100)' }
  ],
  'Date Functions': [
    { name: 'TODAY', syntax: '=TODAY()', description: 'Current date', example: '=TODAY()' },
    { name: 'NOW', syntax: '=NOW()', description: 'Current date and time', example: '=NOW()' },
    { name: 'DATEDIF', syntax: '=DATEDIF(start_date, end_date, "unit")', description: 'Date difference', example: '=DATEDIF(A2, B2, "d")' },
    { name: 'EOMONTH', syntax: '=EOMONTH(start_date, months)', description: 'End of month', example: '=EOMONTH(A2, 0)' }
  ],
  'Dynamic Arrays': [
    { name: 'FILTER', syntax: '=FILTER(array, include, [if_empty])', description: 'Filter data dynamically', example: '=FILTER(A2:A10, B2:B10="Yes")' },
    { name: 'SORT', syntax: '=SORT(array, [sort_index], [sort_order])', description: 'Sort data dynamically', example: '=SORT(A2:A10, 1, -1)' },
    { name: 'UNIQUE', syntax: '=UNIQUE(array, [by_col], [exactly_once])', description: 'Remove duplicates', example: '=UNIQUE(A2:A20)' },
    { name: 'SEQUENCE', syntax: '=SEQUENCE(rows, [columns], [start], [step])', description: 'Generate sequences', example: '=SEQUENCE(10, 1, 1, 1)' }
  ]
};

// Excel Tips and Tricks
const excelTips = [
  {
    category: 'Navigation',
    tips: [
      'Use Ctrl + Arrow keys to jump to the edge of data',
      'Ctrl + Home takes you to cell A1',
      'F5 opens Go To dialog for quick navigation',
      'Use Ctrl + Page Up/Down to switch between sheets'
    ]
  },
  {
    category: 'Selection',
    tips: [
      'Ctrl + A selects all data in current region',
      'Shift + Arrow keys extends selection',
      'Ctrl + Shift + Arrow keys selects to edge of data',
      'Double-click cell border to select to edge'
    ]
  },
  {
    category: 'Editing',
    tips: [
      'F2 enters edit mode for current cell',
      'Ctrl + Enter fills selected cells with same value',
      'Alt + Enter creates new line within cell',
      'Use Ctrl + D to copy from cell above'
    ]
  },
  {
    category: 'Formatting',
    tips: [
      'Ctrl + 1 opens Format Cells dialog',
      'Ctrl + B, I, U for bold, italic, underline',
      'Alt + H + O + I auto-fits column width',
      'Use Format Painter (Ctrl + Shift + C) to copy formatting'
    ]
  },
  {
    category: 'Formulas',
    tips: [
      'F4 cycles through absolute/relative references',
      'Ctrl + ` toggles formula view',
      'Use F9 to evaluate parts of formulas',
      'Ctrl + Shift + Enter for array formulas (legacy)'
    ]
  }
];

/* Render helpers */
function renderResources(list) {
  const grid = $('#resource-grid');
  const empty = $('#resource-empty');
  grid.innerHTML = '';
  if (!list.length) {
    empty.classList.remove('hidden');
    return;
  }
  empty.classList.add('hidden');
  const fragment = document.createDocumentFragment();
  list.forEach(item => {
    const card = document.createElement('article');
    card.className = 'card';
    card.innerHTML = `
      <div class="card-header">
        <h3>${item.title}</h3>
        <span class="difficulty-badge ${item.difficulty.toLowerCase()}">${item.difficulty}</span>
      </div>
      <p>${item.description}</p>
      <div class="card-meta">
        <span class="type-badge">${item.type}</span>
        <span class="duration">${item.duration}</span>
      </div>
      <div class="card-tags">
        ${item.tags.map(tag => `<span class="tag">${tag}</span>`).join('')}
      </div>
      <button class="card-cta" onclick="openResource('${item.title}')">Learn More →</button>
    `;
    fragment.appendChild(card);
  });
  grid.appendChild(fragment);
}

// Enhanced resource viewer with detailed information
function openResource(title) {
  const resource = resources.find(r => r.title === title);
  if (!resource) return;
  
  const modal = document.createElement('div');
  modal.className = 'modal';
  modal.innerHTML = `
    <div class="modal-content">
      <div class="modal-header">
        <h2>${resource.title}</h2>
        <button class="modal-close" onclick="closeModal()">&times;</button>
      </div>
      <div class="modal-body">
        <div class="resource-info">
          <p><strong>Type:</strong> ${resource.type}</p>
          <p><strong>Difficulty:</strong> ${resource.difficulty}</p>
          <p><strong>Duration:</strong> ${resource.duration}</p>
          <p><strong>Description:</strong> ${resource.description}</p>
        </div>
        
        ${resource.detailedInfo ? renderDetailedInfo(resource.detailedInfo) : ''}
        ${resource.tags.includes('Formulas') ? renderFormulaExamples(resource.title) : ''}
        ${resource.tags.includes('Beginner') ? renderExcelTips() : ''}
      </div>
    </div>
  `;
  document.body.appendChild(modal);
  setTimeout(() => modal.classList.add('show'), 10);
}

function renderDetailedInfo(detailedInfo) {
  return `
    <div class="detailed-info">
      <div class="info-section">
        <h3>What You'll Learn</h3>
        <ul>
          ${detailedInfo.whatYouLearn.map(item => `<li>${item}</li>`).join('')}
        </ul>
      </div>
      
      <div class="info-section">
        <h3>Prerequisites</h3>
        <p>${detailedInfo.prerequisites}</p>
      </div>
      
      <div class="info-section">
        <h3>Includes</h3>
        <ul>
          ${detailedInfo.includes.map(item => `<li>${item}</li>`).join('')}
        </ul>
      </div>
      
      <div class="info-section">
        <h3>Learning Outcomes</h3>
        <p>${detailedInfo.outcomes}</p>
      </div>
    </div>
  `;
}

function closeModal() {
  const modal = document.querySelector('.modal');
  if (modal) {
    modal.classList.remove('show');
    setTimeout(() => modal.remove(), 300);
  }
}

function renderFormulaExamples(resourceTitle) {
  let relevantFormulas = [];
  
  // Match resource title to formula categories
  if (resourceTitle.includes('Basic') || resourceTitle.includes('Starter')) {
    relevantFormulas = formulaExamples['Basic Math'];
  } else if (resourceTitle.includes('Text')) {
    relevantFormulas = formulaExamples['Text Functions'];
  } else if (resourceTitle.includes('VLOOKUP') || resourceTitle.includes('Lookup')) {
    relevantFormulas = formulaExamples['Lookup Functions'];
  } else if (resourceTitle.includes('IF') || resourceTitle.includes('Logic')) {
    relevantFormulas = formulaExamples['Logical Functions'];
  } else if (resourceTitle.includes('Date') || resourceTitle.includes('Time')) {
    relevantFormulas = formulaExamples['Date Functions'];
  } else if (resourceTitle.includes('Dynamic') || resourceTitle.includes('Array')) {
    relevantFormulas = formulaExamples['Dynamic Arrays'];
  }
  
  if (relevantFormulas.length === 0) return '';
  
  return `
    <div class="formula-examples">
      <h3>Related Formulas</h3>
      <div class="formula-grid">
        ${relevantFormulas.map(formula => `
          <div class="formula-card">
            <h4>${formula.name}</h4>
            <div class="formula-syntax">${formula.syntax}</div>
            <p>${formula.description}</p>
            <div class="formula-example">
              <strong>Example:</strong> ${formula.example}
            </div>
          </div>
        `).join('')}
      </div>
    </div>
  `;
}

function renderExcelTips() {
  return `
    <div class="excel-tips">
      <h3>Quick Excel Tips</h3>
      <div class="tips-grid">
        ${excelTips.map(category => `
          <div class="tip-category">
            <h4>${category.category}</h4>
            <ul>
              ${category.tips.map(tip => `<li>${tip}</li>`).join('')}
            </ul>
          </div>
        `).join('')}
      </div>
    </div>
    `;
}



/* Enhanced filters and search */
function applyFilters() {
  const activeChip = $('.chip.is-active');
  const tag = activeChip ? activeChip.dataset.tag : 'All';
  const q = ($('#search').value || '').trim().toLowerCase();
  const difficulty = $('#difficulty-filter').value;
  
  const filtered = resources.filter(r => {
    const matchesTag = tag === 'All' || r.tags.includes(tag);
    const hay = `${r.title} ${r.description} ${r.type} ${r.tags.join(' ')}`.toLowerCase();
    const matchesQuery = !q || hay.includes(q);
    const matchesDifficulty = difficulty === 'All' || r.difficulty === difficulty;
    return matchesTag && matchesQuery && matchesDifficulty;
  });
  renderResources(filtered);
}

function setupFilters() {
  const chips = $$('.chip');
  chips.forEach(chip => {
    chip.addEventListener('click', () => {
      chips.forEach(c => c.classList.remove('is-active'));
      chip.classList.add('is-active');
      chips.forEach(c => c.setAttribute('aria-pressed', String(c === chip)));
      applyFilters();
    });
  });
  $('#search').addEventListener('input', applyFilters);
  $('#difficulty-filter').addEventListener('change', applyFilters);
}



/* Newsletter form */
function setupNewsletter() {
  const form = $('#newsletter-form');
  const email = $('#email');
  const message = $('#newsletter-message');
  form.addEventListener('submit', (e) => {
    e.preventDefault();
    message.textContent = '';
    const value = email.value.trim();
    const ok = /^\S+@\S+\.\S+$/.test(value);
    if (!ok) {
      message.textContent = 'Please enter a valid email address.';
      message.style.color = '#ff9b9b';
      email.focus();
      return;
    }
    form.reset();
    message.textContent = 'Thanks! Please check your inbox to confirm.';
    message.style.color = '';
  });
}

/* KPI count-up */
function animateKpis() {
  const map = [
    { el: $('#kpi-tutorials'), target: 120, suffix: '+' },
    { el: $('#kpi-templates'), target: 45, suffix: '+' },
    { el: $('#kpi-commands'), target: 25, suffix: '+' },
    { el: $('#kpi-members'), target: 5000, suffix: '+' },
  ];
  const observer = new IntersectionObserver((entries, obs) => {
    entries.forEach(entry => {
      if (!entry.isIntersecting) return;
      map.forEach(({ el, target, suffix }) => {
        const start = 0;
        const duration = 1200;
        const startTs = performance.now();
        const formatter = new Intl.NumberFormat(undefined);
        function tick(now) {
          const t = Math.min(1, (now - startTs) / duration);
          const eased = 1 - Math.pow(1 - t, 3);
          const val = Math.round(start + (target - start) * eased);
          el.textContent = formatter.format(val) + suffix;
          if (t < 1) requestAnimationFrame(tick);
        }
        requestAnimationFrame(tick);
      });
      obs.disconnect();
    });
  }, { threshold: 0.3 });
  observer.observe($('.kpis'));
}

/* Footer year */
function setYear() {
  const el = $('#year');
  if (el) el.textContent = new Date().getFullYear();
}

/* Run */
document.addEventListener('DOMContentLoaded', () => {
  renderResources(resources);
  setupFilters();
  setupNewsletter();
  animateKpis();
  setYear();
});


