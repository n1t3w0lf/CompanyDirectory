import * as React from 'react';
import { Stack } from '@fluentui/react/lib/Stack';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { TextField } from '@fluentui/react/lib/TextField';
import { Separator } from '@fluentui/react/lib/Separator';

export interface IAdvancedFilters {
  department?: string;
  officeLocation?: string;
  city?: string;
  country?: string;
  jobTitle?: string;
}

export interface IAdvancedFilterPanelProps {
  departments: string[];
  locations: string[];
  cities: string[];
  countries: string[];
  jobTitles: string[];
  currentFilters: IAdvancedFilters;
  onApplyFilters: (filters: IAdvancedFilters) => void;
  onClearFilters: () => void;
}

export const AdvancedFilterPanel: React.FC<IAdvancedFilterPanelProps> = ({
  departments,
  locations,
  cities,
  countries,
  jobTitles: _jobTitles,
  currentFilters,
  onApplyFilters,
  onClearFilters
}) => {
  const [filters, setFilters] = React.useState<IAdvancedFilters>(currentFilters);

  const departmentOptions: IDropdownOption[] = React.useMemo(() => [
    { key: '', text: 'All Departments' },
    ...departments.map(d => ({ key: d, text: d }))
  ], [departments]);

  const locationOptions: IDropdownOption[] = React.useMemo(() => [
    { key: '', text: 'All Locations' },
    ...locations.map(l => ({ key: l, text: l }))
  ], [locations]);

  const cityOptions: IDropdownOption[] = React.useMemo(() => [
    { key: '', text: 'All Cities' },
    ...cities.map(c => ({ key: c, text: c }))
  ], [cities]);

  const countryOptions: IDropdownOption[] = React.useMemo(() => [
    { key: '', text: 'All Countries' },
    ...countries.map(c => ({ key: c, text: c }))
  ], [countries]);

  const handleDepartmentChange = React.useCallback((_: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
    setFilters(prev => ({ ...prev, department: option?.key as string || undefined }));
  }, []);

  const handleLocationChange = React.useCallback((_: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
    setFilters(prev => ({ ...prev, officeLocation: option?.key as string || undefined }));
  }, []);

  const handleCityChange = React.useCallback((_: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
    setFilters(prev => ({ ...prev, city: option?.key as string || undefined }));
  }, []);

  const handleCountryChange = React.useCallback((_: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
    setFilters(prev => ({ ...prev, country: option?.key as string || undefined }));
  }, []);

  const handleJobTitleChange = React.useCallback((_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
    setFilters(prev => ({ ...prev, jobTitle: newValue || undefined }));
  }, []);

  const handleApply = React.useCallback((): void => {
    onApplyFilters(filters);
  }, [filters, onApplyFilters]);

  const handleClear = React.useCallback((): void => {
    const emptyFilters: IAdvancedFilters = {};
    setFilters(emptyFilters);
    onClearFilters();
  }, [onClearFilters]);

  return (
    <Stack tokens={{ childrenGap: 16 }} styles={{ root: { padding: '20px 0' } }}>
      <Dropdown
        label="Department"
        options={departmentOptions}
        selectedKey={filters.department || ''}
        onChange={handleDepartmentChange}
        placeholder="Select a department"
      />

      <Dropdown
        label="Office Location"
        options={locationOptions}
        selectedKey={filters.officeLocation || ''}
        onChange={handleLocationChange}
        placeholder="Select a location"
      />

      <Dropdown
        label="City"
        options={cityOptions}
        selectedKey={filters.city || ''}
        onChange={handleCityChange}
        placeholder="Select a city"
      />

      <Dropdown
        label="Country"
        options={countryOptions}
        selectedKey={filters.country || ''}
        onChange={handleCountryChange}
        placeholder="Select a country"
      />

      <TextField
        label="Job Title (search)"
        value={filters.jobTitle || ''}
        onChange={handleJobTitleChange}
        placeholder="Enter job title keywords"
      />

      <Separator />

      <Stack horizontal tokens={{ childrenGap: 8 }}>
        <PrimaryButton onClick={handleApply} text="Apply Filters" />
        <DefaultButton onClick={handleClear} text="Clear All" />
      </Stack>
    </Stack>
  );
};
