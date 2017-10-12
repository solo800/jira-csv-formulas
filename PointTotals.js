function PointTotals (range, teamMembers) {
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
  this._init = function (range, teamMembers) {
    this.DAY_MILLISECONDS = 86400000;
    this.rawRange = range.slice(); // Copy array
    this.range = range;
    this._setHeader();
    this._setHeaderKeys();

    // Remove the header row from the range
    this.range.splice(0, 1);

    this.teamMembers = undefined !== teamMembers ? teamMembers[0] : this._setTeamMembers();
    // Remove blanks from teamMembers
    this.teamMembers = this.teamMembers.filter(function (member) {
      return '' !== member;
    });

    this.range = this.range.filter(function (row) {
      return !isNaN(parseInt(this.getEstimate(row)));
    }, this);

    this.chunkRange(function () {
      this.chunks = [['Team Member', 'Points']];
      var getEstimate = this.getEstimate;

      this.teamMembers.forEach(function (member) {
        this.chunks.push([member,
          this.range.filter(function (row) {
            return member === row[this.headerKeys.assignee];
          }, this).reduce(function (acc, row) {
            return acc + getEstimate(row);
          }, 0)
        ]);
      }, this);
    });
  };

  return this._init(range, teamMembers);
}
