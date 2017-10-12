function TicketsByComplexity (range, DURATION, teamMembers) {
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
  this.outputChunks = function (columns) {
    var output = [columns];
    var outputChunk;

    this.processedChunks.forEach(function (chunk, i) {
      outputChunk = [new Date(this.chunks[i].start)];

      chunk.forEach(function (points) {
        outputChunk.push(points);
      });
      output.push(['week of ' + this._formatDate(new Date(this.chunks[i].start))].concat(chunk));
      // output.push(['week of ' + this._formatDate(new Date(this.chunks[i].start))].concat(chunk));
    }, this);

    return output;
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
  this.chunkRange = function (cb) {
    cb.call(this);
  };
  this.processChunks = function (cb) {
    cb.call(this);
  }
  this.getEstimate = function (row) {
    var orig = parseInt(row[this.headerKeys['original estimate']]);
    var score = parseInt(row[this.headerKeys['sprint score']]);
    return !isNaN(orig) ? orig / 3600 : score;
  }
  // Private methods
  this._setTeamMembers = function () {
    this.teamMembers = this.range.map(function (row) {
      return row[this.headerKeys.assignee];
    }, this).reduce(function (acc, assignee) {
      return -1 === acc.indexOf(assignee) ? acc.concat([assignee]) : acc;
    }, []);

    return this.teamMembers;
  };
  this._setHeader = function () {
    this.header = this.range[0].map(function (label) {
      return label.replace(/custom field|\(|\)/ig, '').trim().toLowerCase();
    });

    return this.header;
  }
  this._dateTimeToDate = function (date) {
    date = date instanceof Date ? date : new Date(date);
    return new Date(date.getFullYear(), date.getMonth(), date.getDate());
  }
  this._formatDate = function (date) {
    return (date.getMonth() + 1) + '/' + date.getDate() + '/' + date.getFullYear();
  }
  this._setHeaderKeys = function () {
    this.headerKeys = {};

    this.header.forEach(function (label, i) {
      this.headerKeys[label] = i;
    }, this);

    return this.headerKeys;
  }
  // Init
  this._init = function (range, DURATION, teamMembers) {
    this.DURATION = undefined !== DURATION ? DURATION : 7;
    this.teamMembers = undefined !== teamMembers ? teamMembers : ['bmarshall', 'jslavin', 'slakshmanan', 'jadams', 'nathan', 'asolomon']
    this.DAY_MILLISECONDS = 86400000;
    this.rawRange = range.slice(); // Copy array
    this.header = this._formatHeader(range);
    this.range = range;

    // Setup a header key so that we can easily reference the correct index for a column when using only the label
    var headerKeys = {};
    this.header.forEach(function (label, i) {
      headerKeys[label] = i;
    });
    this.headerKeys = headerKeys;

    this.range = this.range.filter(function (row) {
      // Filter out:
      // blank rows
      // rows with no estimate
      // rows with an assignee that is not in team members
      return !(
        0 === row.length ||
        0 === this.getEstimate(row) ||
        -1 === this.teamMembers.indexOf(row[this.headerKeys.assignee])
      );
    });

    this.orderBy('resolved');

    this.chunkRange(function () {
      this.chunks = [];
      var colIndex = this.headerKeys.resolved;
      var key;
      var found;
      var toDate = this._dateTimeToDate;

      // Set up the chunks
      var start = this.range[0][colIndex].getTime();
      var end = start + (this.DAY_MILLISECONDS * this.DURATION);
      var limit = this.range[this.range.length - 1][colIndex].getTime();

      var li = 1000;
      while (start < limit && li > 0) {
        li--;
        this.chunks.push({
          start: start,
          end: end,
          rows: [],
        });

        start = end;
        end = start + (this.DAY_MILLISECONDS * this.DURATION);
      }

      this.range.forEach(function (row) {
        if (row[colIndex] instanceof Date) {
          // Iterate over all chunks checking to see if this fits in one
          this.chunks.forEach(function (chunk) {
            if (chunk.start <= row[colIndex].getTime() && chunk.end > row[colIndex].getTime()) {
              chunk.rows.push(row);
            }
          });
        }
      }, this);
    });

    this.processChunks(function () {
      this.processedChunks = [];
      var chunkRow;
      var points;

      this.chunks.forEach(function (chunk) {
        chunkRow = [];
        // Grab all rows for this team member and add them to the range
        this.teamMembers.forEach(function (assignee) {
          points = 0;

          chunk.rows.forEach(function (row) {
            if (assignee === row[this.headerKeys.assignee]) {
              points += this.getEstimate(row);
            }
          }, this);

          chunkRow.push(points);
        }, this);

        this.processedChunks.push(chunkRow);
      }, this);
    });

    return this.outputChunks(['week of '].concat(this.teamMembers));
  }
  return this._init(range, DURATION, teamMembers);
}
