import * as React from 'react';
import styles from './GroupPeople.module.scss';
import { IGroupPeopleProps } from './IGroupPeopleProps';

import { Persona } from 'office-ui-fabric-react/lib/Persona';
import PeopleCard from '../models/PeopleCard';
import { DefaultButton, PrimaryButton, Stack, IStackTokens } from 'office-ui-fabric-react';

import * as strings from 'GroupPeopleWebPartStrings';

export interface IStateGroupPeople {
    _viewAll: boolean;
}

/** Group People UI
 * @class
 * @extends
 */
export default class GroupPeople extends React.Component<IGroupPeopleProps, IStateGroupPeople> {

    private _aignment: string = '';

    /** Toggle Title state
     * @private
     */
    private _toggleTitle: string = '';

    /** Display a message if no People to display and if don't hide webpart
     * @private
     */
    private _displayDefaultMessage: string = '';

    /** Default constructor
     * @param props 
     */

    

    constructor(props: IGroupPeopleProps) {
        super(props);
        this._toggleTitle = props.displayTitle ? '' : styles.hidden;
        this.state = {
            _viewAll: false
        };
    }

    public viewAll() {
        console.log("View All clicked");
        this.setState((prevState: IStateGroupPeople) => ({
            _viewAll: !prevState._viewAll
        })); 
    }

    /** Default render
     * @returns HTML Template
     * @public
     */
    public render(): JSX.Element {
        console.log("Richen render -->", this.props.numberOfUser);
        this._toggleTitle = this.props.displayTitle ? '' : styles.hidden;
        this._displayDefaultMessage = (this.props.users.length == 0 && !this.props.hide) ? '' : styles.hidden;
        this._aignment = (true) ? styles.vertical : styles.horizontal;
        return (
            <div className={styles.groupPeople}>
                <div className={styles.container}>
                    <div className={styles.row}>
                        <div className={styles.column}>
                            <h2 className={[styles.title, this._toggleTitle].join(' ')} role="heading">{this.props.title}</h2>
                            <div className={['grpPeopleNoItem', this._displayDefaultMessage].join(' ')}>{strings.NoItemFound}</div>
                            <div className={this._aignment}>
                                {this.props.users.slice(0, (this.state._viewAll) ? this.props.users.length : this.props.numberOfUser).map((p: PeopleCard) => {
                                    return (<div className={styles.personaTile} key={p.key}><Persona
                                        text={p.lineOne}
                                        secondaryText={p.lineTwo}
                                        tertiaryText={p.lineThree}
                                        imageUrl={p.image}
                                        size={this.props.size}
                                        className={styles.persona}
                                    /></div>);
                                })}
                            </div>
                            {
                                (!this.state._viewAll && this.props.users.length > this.props.numberOfUser) ? <DefaultButton text="View All" onClick={() => { this.viewAll() }} /> : <div></div>
                            }
                        </div>
                    </div>
                </div>
            </div>
        );
    }
}
