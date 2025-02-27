const fs = require('fs');
const XLSX = require('xlsx');
const crypto = require('crypto');

// Define the hierarchy levels globally so they can be accessed in callbacks
const hierarchyLevels = [
  'L1 - Line of Business',
  'L2 - Process Group',
  'L3 - Scope Item',
  'L4 - Process Variant',
  'L5 - Process Step'
];

// Function to generate a MongoDB-style ObjectId
function generateObjectId() {
  return crypto.randomBytes(12).toString('hex');
}

// Main function to process the Excel file
async function processExcelToTreeJSON(inputFilePath, outputFilePath) {
  try {
    console.log(`Reading Excel file: ${inputFilePath}`);
    const excelData = fs.readFileSync(inputFilePath);
    
    // Parse the Excel file
    const workbook = XLSX.read(excelData, {
      cellStyles: true,
      cellFormulas: true,
      cellDates: true,
      cellNF: true,
      sheetStubs: true
    });
    
    // Get the first sheet
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    
    // Convert to JSON objects with headers
    const rawData = XLSX.utils.sheet_to_json(worksheet);
    
    console.log(`Processing ${rawData.length} records...`);
    
    // Output array will store all node objects
    const outputNodes = [];
    
    // Map to store generated MongoDB IDs for each unique combination of level+value
    // This ensures we create separate entries for the same value at different levels
    const idMap = new Map();
    
    // Map to quickly look up nodes by their ID
    const nodeMap = new Map();
    
    // Process each row in the Excel data
    rawData.forEach(row => {
      let previousLevelId = null;
      
      // Process each level in this row
      hierarchyLevels.forEach((level, levelIndex) => {
        if (!row[level] || row[level].trim() === '') {
          return; // Skip empty values
        }
        
        const value = row[level].trim();
        
        // Create a composite key that includes the parent info to handle the same value under different parents
        // For L1, there's no parent, so just use the level and value
        const compositeKey = previousLevelId 
          ? `${level}:${value}:${previousLevelId}` 
          : `${level}:${value}`;
        
        // Check if we've already created this node
        if (!idMap.has(compositeKey)) {
          // Generate a new MongoDB-style ID
          const nodeId = generateObjectId();
          
          // Create the node object
          const node = {
            _id: nodeId,
            title: value,
            _parent: previousLevelId || '',
            _child: [],
            // Set all hierarchy fields to empty except the current level
            ...Object.fromEntries(hierarchyLevels.map(l => [l, l === level ? value : ''])),
            // Include metadata fields from the row
            ID: row.ID || '',
            'Business Role': row['Business Role'] || '',
            'Fiori app UX recommendations': row['Fiori app UX recommendations'] || '',
            'Insights (Indicative)': row['Insights (Indicative)'] || '',
            'Business stakeholders': row['Business stakeholders'] || '',
            'Materiality': row['Materiality'] || 0,
            'Description': row['Description'] || ''
          };
          
          // Add to output array
          outputNodes.push(node);
          
          // Store the ID in our maps
          idMap.set(compositeKey, nodeId);
          nodeMap.set(nodeId, node);
          
          // If this node has a parent, add this node as a child to the parent
          if (previousLevelId) {
            const parentNode = nodeMap.get(previousLevelId);
            if (parentNode && !parentNode._child.includes(nodeId)) {
              parentNode._child.push(nodeId);
            }
          }
        }
        
        // Update previousLevelId for the next iteration
        previousLevelId = idMap.get(compositeKey);
      });
    });
    
    // Write the output JSON file
    fs.writeFileSync(outputFilePath, JSON.stringify(outputNodes, null, 2));
    
    console.log(`Successfully processed ${outputNodes.length} nodes`);
    console.log(`Output written to: ${outputFilePath}`);
    
    return outputNodes;
  } catch (error) {
    console.error('Error processing Excel file:', error);
    throw error;
  }
}

// Command line arguments handling
const inputFile = process.argv[2] || 'digitalmapsresults_scopeitem_metawarsss4i1_s4hana_onprem_1909.xlsx';
const outputFile = process.argv[3] || 'tree_output.json';

// Verify that input file exists
if (!fs.existsSync(inputFile)) {
  console.error(`Error: Input file '${inputFile}' does not exist.`);
  console.log('Usage: node script.js [inputFile.xlsx] [outputFile.json]');
  process.exit(1);
}

// Run the process
processExcelToTreeJSON(inputFile, outputFile)
  .then(() => {
    console.log('Process completed successfully');
    
    try {
      // Read the output file to display some statistics
      const output = JSON.parse(fs.readFileSync(outputFile));
      
      // Count nodes at each level
      const levelCounts = {};
      hierarchyLevels.forEach(level => {
        levelCounts[level] = output.filter(node => node[level] !== '').length;
      });
      
      console.log('\nNodes at each level:');
      Object.entries(levelCounts).forEach(([level, count]) => {
        console.log(`- ${level}: ${count} nodes`);
      });
      
      // Find items that appear multiple times with different parents
      const titleCounts = {};
      output.forEach(node => {
        if (!titleCounts[node.title]) {
          titleCounts[node.title] = [];
        }
        titleCounts[node.title].push(node._id);
      });
      
      // Check for duplicated items
      const duplicatedItems = Object.entries(titleCounts)
        .filter(([title, ids]) => ids.length > 1)
        .sort((a, b) => b[1].length - a[1].length);
      
      if (duplicatedItems.length > 0) {
        console.log(`\nFound ${duplicatedItems.length} items that appear under multiple parents:`);
        
        // Display the top 5 most duplicated items
        duplicatedItems.slice(0, 5).forEach(([title, ids]) => {
          console.log(`- "${title}" appears ${ids.length} times with different parents`);
          
          // Show a few examples
          const examples = output.filter(node => node.title === title).slice(0, 3);
          examples.forEach(node => {
            const parent = output.find(n => n._id === node._parent);
            if (parent) {
              const parentLevel = hierarchyLevels.find(level => parent[level] !== '');
              const nodeLevel = hierarchyLevels.find(level => node[level] !== '');
              console.log(`  * ${nodeLevel}: "${node.title}" under ${parentLevel}: "${parent.title}"`);
            }
          });
          
          if (ids.length > 3) {
            console.log(`  * ... and ${ids.length - 3} more instances`);
          }
        });
      }
    } catch (err) {
      console.error('Error analyzing output:', err);
    }
  })
  .catch(err => console.error('Process failed:', err));