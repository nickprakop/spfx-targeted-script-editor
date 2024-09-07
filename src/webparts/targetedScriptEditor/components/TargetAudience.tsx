import { PageContext } from "@microsoft/sp-page-context";
import { IPropertyFieldGroupOrPerson } from "@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker";
import * as React from "react";
import spservices from "../../../services/spservices";
export interface ITargetAudienceProps {
    pageContext: PageContext;
    groupIds: IPropertyFieldGroupOrPerson[];
}
export interface ITargetAudienceState {
    canView?: boolean;
}
export default class TargetAudience extends React.Component<ITargetAudienceProps, ITargetAudienceState>{
    constructor(props: ITargetAudienceProps) {
        super(props);
        this.state = {
            canView: false
        } as ITargetAudienceState;

    }
    public componentDidMount(): void {
        //setting the state whether user has permission to view webpart
        this.checkUserCanViewWebpart();
    }
    public render(): JSX.Element {
        return (
            <div>
                {
                    this.props.groupIds
                        ? (
                            this.state.canView
                                ? this.props.children
                                : ``
                        )
                        : this.props.children
                }
            </div>
        );
    }
    public checkUserCanViewWebpart(): void {
        const self = this;
        let proms: Promise<any>[] = [];
        const errors: any[] = [];
        const _sv = new spservices();
        self.props.groupIds.map((item) => {
            proms.push(_sv.isMember(item.fullName, self.props.pageContext.legacyPageContext[`userId`], self.props.pageContext.site.absoluteUrl));
        });
        Promise.race(
            proms.map(p => {
                return p.catch(err => {
                    errors.push(err);
                    if (errors.length >= proms.length) throw errors;
                    // Ignore the error and return a new promise that never settles
                    return new Promise(() => { });
                });
            })).then(val => {
                // This will only run if one of the promises succeeds
                this.setState({ canView: true }); 
            });
    }
}
