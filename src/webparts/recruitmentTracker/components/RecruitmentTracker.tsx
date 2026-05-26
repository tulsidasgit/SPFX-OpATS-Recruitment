import * as React from 'react';
import {
  Pivot,
  PivotItem,
  Spinner,
  SpinnerSize,
  Stack,
  Text,
} from '@fluentui/react';
import type { IRecruitmentTrackerProps } from './IRecruitmentTrackerProps';
import { PostJobForm } from './PostJob/PostJobForm';
import { TrackProgress } from './TrackProgress/TrackProgress';

interface IRecruitmentTrackerState {
  activeTab: string;
}

export default class RecruitmentTracker extends React.Component<
  IRecruitmentTrackerProps,
  IRecruitmentTrackerState
> {
  constructor(props: IRecruitmentTrackerProps) {
    super(props);
    this.state = { activeTab: 'postJob' };
  }

  public render(): React.ReactElement<IRecruitmentTrackerProps> {
    const { sp, graphService, currentUser, applyBaseUrl } = this.props;

    if (!currentUser || !currentUser.email) {
      return (
        <Stack horizontalAlign="center" verticalAlign="center" styles={{ root: { height: 200 } }}>
          <Spinner size={SpinnerSize.large} label="Loading Recruitment Tracker…" />
        </Stack>
      );
    }

    return (
      <Stack tokens={{ childrenGap: 0 }}>
        {/* Header bar */}
        <Stack
          horizontal
          horizontalAlign="space-between"
          verticalAlign="center"
          styles={{
            root: {
              background: '#fd800b',
              padding: '12px 24px',
            },
          }}
        >
          <Text styles={{ root: { color: '#fff', fontSize: 18, fontWeight: 600 } }}>
            OpATS — Recruitment Tracker
          </Text>
          <Text styles={{ root: { color: '#fff', fontSize: 13 } }}>
            {currentUser.displayName}
          </Text>
        </Stack>

        {/* Tab navigation */}
        <Pivot
          selectedKey={this.state.activeTab}
          onLinkClick={item => item && this.setState({ activeTab: item.props.itemKey ?? 'postJob' })}
          styles={{
            root: { paddingLeft: 16 },
            link: { selectors: { '&.ms-Pivot-link.is-selected': { borderBottomColor: '#fd800b', color: '#fd800b' } } },
          }}
        >
          <PivotItem headerText="Post Job" itemKey="postJob" itemIcon="Add">
            <PostJobForm
              sp={sp}
              graphService={graphService}
              currentUser={currentUser}
              applyBaseUrl={applyBaseUrl}
            />
          </PivotItem>

          <PivotItem headerText="Track Progress" itemKey="trackProgress" itemIcon="ClipboardList">
            <TrackProgress sp={sp} graphService={graphService} />
          </PivotItem>

          <PivotItem headerText="Completed Jobs" itemKey="completedJobs" itemIcon="CheckMark">
            <Stack tokens={{ padding: 24 }}>
              <Text variant="large">Completed Jobs</Text>
              <Text styles={{ root: { color: '#605e5c' } }}>
                This tab is coming in Phase 11.
              </Text>
            </Stack>
          </PivotItem>
        </Pivot>
      </Stack>
    );
  }
}
