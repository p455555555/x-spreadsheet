// import helper from '../helper';

export default class History {
  constructor() {
    this.undoItems = [];
    this.redoItems = [];
  }

  add(data) {
    console.log('History add>>>>>', data);
    this.undoItems.push(JSON.stringify(data));
    this.redoItems = [];
  }

  canUndo() {
    return this.undoItems.length > 0;
  }

  canRedo() {
    return this.redoItems.length > 0;
  }

  undo(rows) {
    const { undoItems, redoItems } = this;
    if (this.canUndo()) {
      const popData = JSON.parse(undoItems.pop());
      const nowData = this.getRowNowData(rows, popData);
      redoItems.push(nowData);
      console.log('redoItems>>>>>>>', redoItems);
      console.log('undoItems>>>>>>>', undoItems);
      return popData;
    }
    return false;
  }

  redo(rows) {
    const { undoItems, redoItems } = this;
    if (this.canRedo()) {
      const popData = JSON.parse(redoItems.pop());
      undoItems.push(this.getRowNowData(rows, popData));
      console.log('redoItems>>>>>>>', redoItems);
      console.log('undoItems>>>>>>>', undoItems);
      return popData;
    }
    return false;
  }

  getRowNowData(rows, data) {
    const nowData = {};
    for (const key in data) {
      if (Object.hasOwnProperty.call(data, key)) {
        nowData[key] = rows._[key] || null;
      }
    }

    return JSON.stringify(nowData);
  }
}
