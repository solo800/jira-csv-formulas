function PointsOverSprint (range, DURATION) { 
  /**
   * Orders by column
   * if column is number it is assumed to be an index of range to order by
   * if column is a string it is assumed to be a label within header
   */
  this.orderBy = function (column, ascending, sortCb) {
    var colIndex = 'number' === column ? colIndex = column : this.header.indexOf(column);
    ascending = undefined === ascending ? true : false;
    sortCb = 'function' === typeof sortCb ? sortCb : null;
    
    this.range.sort('function' === typeof sortCb ? sortCb : function (prev, next) {
      if (prev[colIndex] instanceof Date) {
        var prevTime = prev[colIndex].getTime();
        var nextTime = next[colIndex].getTime();
        return true === ascending ? prevTime - nextTime : nextTime - prevTime;
      } else if (!isNaN(parseInt(prev[colIndex]))) {
        var prevNum = 1 * prev[colIndex];
        var nextNum = 1 * next[colIndex];
        return true === ascending ? prevNum - nextNum : nextNum - prevNum;
      } else {
        var prevStr = prev[colIndex];
        var nextStr = next[colIndex];
        return true === ascending ? prevStr > nextStr : nextStr > prevStr;
      }
    });
  };
  this.columnsToDate = function (columns) {
    var header = this.header;
    var range = this.range;
    var colIndex;
    var tempDate;

    columns.forEach(function (column) {
      colIndex = 'number' === typeof column ? column : header.indexOf(column);
      
      range = range.map(function (row) {
        if (!(row[colIndex] instanceof Date)) {
          tempDate = new Date(row[colIndex]);
          
          if (!isNaN(tempDate.getTime())) {
            row[colIndex] = tempDate;
          }
        }
        return row;
      });
    });
    
    this.range = range;
  };
  this.filterRow = function (cb) {
    this.range = this.range.filter(cb, this);
  };
  this.chunkRange = function (cb) {
    cb.call(this);
  };
  this.processChunks = function (cb) {
    cb.call(this);
  };
  this.getEstimate = function (row) {
    var orig = parseInt(row[this.headerKeys['original estimate']]);
    var score = parseInt(row[this.headerKeys['sprint score']]);
    return !isNaN(orig) ? orig / 3600 : score;
  };
  // Only works on columns with strings as values
  this.simplifyColumns = function (columnLabel) {
    columnLabel = undefined === columnLabel ? 'sprint' : columnLabel;
    
    // Flaten the Sprint columns
    this.range = this.range.map(function (row, i, range) {
      var values = [];

      // Put all the sprints into an array
      this.header.forEach(function (label, i) {
        if (columnLabel === label) {
          if ('' !== row[i]) {
            values.push(row[i]);
          }
        }
      });
      
      // Sort the sprints and grab the last one and add it at the end of the row, that's our new row
      mappedRow = row.filter(function (col, i) {
        return columnLabel !== this.header[i];
      }, this);

      mappedRow.push(values.sort().pop());
      
      return mappedRow;
    }, this);
    
    // Make sure to update the header to have only one sprint column
    this.header = this.header.filter(function (label) {
      return columnLabel !== label;
    });
    
    this.header.push(columnLabel);
    
    this._setHeaderKeys();
  };
  // Private methods
  this._formatHeader = function (range) {
    return range.shift().map(function (label) {
      return label.replace(/custom field|\(|\)/ig, '').trim().toLowerCase();
    });
  };
  this._dateTimeToDate = function (date) {
    date = date instanceof Date ? date : new Date(date);
    return new Date(date.getFullYear(), date.getMonth(), date.getDate());
  };
  this._formatDate = function (date) {
    return date.getMonth() + '/' + date.getDate() + '/' + date.getFullYear();
  };
  this._setHeaderKeys = function () {
    this.headerKeys = {};
    this.header.forEach(function (label, i) {
      this.headerKeys[label] = i;
    }, this);
    return this.headerKeys;
  };
  // Init
  this._init = function (range, DURATION, teamMembers) {
    this.DURATION = undefined !== DURATION ? DURATION : 7;
    this.teamMembers = undefined !== teamMembers ? teamMembers : ['bmarshall', 'jslavin', 'slakshmanan', 'jadams', 'nathan', 'asolomon']
    this.DAY_MILLISECONDS = 86400000;
    this.rawRange = range.slice(); // Copy array
    this.header = this._formatHeader(range);
    this.range = range;
    
    // Setup a header key so that we can easily reference the correct index for a column when using only the label
    this._setHeaderKeys();
    
//    this.columnsToDate(['resolved']);
    
    this.filterRow(function (row) {      
      // Filter out rows without either an estimate or sprint score, no assignee, no resolved
      if (
        (isNaN(parseInt(row[this.headerKeys['original estimate']])) && isNaN(parseInt(row[this.headerKeys['sprint score']]))) || 
        (-1 === this.teamMembers.indexOf(row[this.headerKeys.assignee]))
      ) {
        return false;
      } else {
        return true;
      }
    });

    this.simplifyColumns('sprint');
    
    this.chunkRange(function () {
      this.chunks = [];
      var colIndex = this.headerKeys.sprint;
      var key;
      var found;
      
      this.range.forEach(function (row) {
        // Iterate over all chunks checking to see if this fits in one
        found = false;
        this.chunks.forEach(function (chunk) {
          if (chunk.sprint === row[colIndex]) {
            chunk.rows.push(row);
            found = true;
          }
        }, this);
        
        // If the chunk wasn't found then we have a new one, add a chunk object with a single row
        if (false === found) {
          this.chunks.push({
            sprint: row[colIndex],
            rows: [row],
          });
        }
      }, this);
    });

    this.processChunks(function () {
      function getMemberScoreFromChunkRows (chunk, member) {
        return chunk.rows.filter(function (row) {
          return row[this.headerKeys.assignee] === member;
        }, this).reduce(function (acc, row) {
          return acc + getEstimate(row);
        }, 0);
      }
      
      var sprintNames = [];
      this.chunks.forEach(function (chunk) {
        if (-1 === sprintNames.indexOf(chunk.sprint)) {
          sprintNames.push(chunk.sprint);
        }
      });
      sprintNames.sort();
      
      // Loop through each sprint
      this.processedChunks = [];
      var processedChunk; // Will become the row added to processedChunks
      var chunk;
      var totalEstimate;
      sprintNames.forEach(function (sprintName) {
        processedChunk = [sprintName];
        
        // Get this sprints chunk
        chunk = this.chunks.filter(function (chunk) {
          return sprintName === chunk.sprint;
        }, this)[0];

        // Output each team member's score for this sprint
        this.teamMembers.forEach(function (teamMemberName) {
          // Get all this members rows from current chunk and reduce down to the total estimate
          totalEstimate = getMemberScoreFromChunkRows(chunk, teamMemberName);
        
          processedChunk.push(totalEstimate);
        }, this);

        this.processedChunks.push(processedChunk);
      }, this);
    });
    
    return [['Sprints'].concat(this.teamMembers)].concat(this.processedChunks);
  }
  return this._init(range, DURATION);
}
