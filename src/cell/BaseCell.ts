import { Cell } from 'exceljs';

import { Scope } from '../Scope';

export declare type CellType = typeof BaseCell;

export /* abstract */ class BaseCell {
    // can't be abstract :(
    /**
     * check if this commend can parse `value`
     */
    public static match(cell: Cell): boolean {
        return false;
    }

    constructor() {
        if (this.constructor.name !== 'BaseCell') {
            return;
        }

        // can't be marked by abstract keyword, so it throw type error.
        throw new TypeError(`Cannot construct ${BaseCell.name} instances directly. It's abstract.`);
    }

    public apply(scope: Scope): BaseCell {
        if (scope.isOutOfColLimit()) {
            scope.finish(); // todo important: spec test
        }
        const templateCell = scope.template.worksheets[scope.templateCell.ws].getCell(
            scope.templateCell.r,
            scope.templateCell.c,
        );
        if (
            templateCell &&
            templateCell.isMerged &&
            templateCell.master &&
            templateCell.master.address !== templateCell.address
        ) {
            // this is a MergeSlaveCell
            scope.applyMerge();
            return this;
        }

        scope.setCurrentOutputValue(scope.getCurrentTemplateValue());
        scope.applyStyles();
        // console.log(
        //     'applying merge for outputCell',
        //     scope.outputCell,
        //     'templateCell',
        //     scope.templateCell,
        //     'masters',
        //     scope.masters,
        // );
        scope.applyMerge();

        return this;
    }
}
