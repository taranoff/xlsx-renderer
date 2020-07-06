import { Cell, ValueType } from 'exceljs';

import { BaseCell } from './BaseCell';
import { Scope } from '../Scope';

export class MergedSlaveCell extends BaseCell {
    /**
     * @inheritDoc
     * @param {Cell} cell
     * @returns {boolean}
     */
    public static match(cell: Cell): boolean {
        return cell && cell.isMerged && cell.master && cell.master.address !== cell.address;
    }

    /**
     * @inheritDoc
     * @param {Scope} scope
     * @returns {NormalCell}
     */
    public apply(scope: Scope): MergedSlaveCell {
        // super.apply(scope);

        scope.incrementCol();

        return this;
    }
}
