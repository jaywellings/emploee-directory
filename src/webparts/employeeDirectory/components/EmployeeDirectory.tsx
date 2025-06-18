import * as React from 'react';

import { useState, useMemo, useEffect } from "react"
import {
  SearchBox,
  CommandBar,
  type ICommandBarItemProps,
  DetailsList,
  DetailsListLayoutMode,
  type IColumn,
  SelectionMode,
  Persona,
  PersonaSize,
  PersonaPresence,
  Stack,
  Text,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  DefaultButton,
  Panel,
  PanelType,
  Pivot,
  PivotItem,
  DocumentCard,
  DocumentCardPreview,
  DocumentCardTitle,
  DocumentCardActivity,
  DocumentCardDetails,
  type IDocumentCardPreviewProps,
  ImageFit,
} from "@fluentui/react"
import type { SharePointUser } from "../types/types"
import { sharePointService } from '../service/dataService';

interface FluentEmployeeDirectoryProps {
  listName?: string
}

export const EmployeeDirectory: React.FC<FluentEmployeeDirectoryProps> = ({ listName }) => {
  const [searchTerm, setSearchTerm] = useState("")
  const [selectedDepartment, setSelectedDepartment] = useState<string>("all")
  const [viewMode, setViewMode] = useState<"list" | "cards" | "people">("cards")
  const [selectedEmployee, setSelectedEmployee] = useState<SharePointUser | null>(null)
	const [isPanelOpen, setIsPanelOpen] = useState(false)
	const [employees, setEmployees] = useState<SharePointUser[]>([])
	const [loading, setLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)


	async function loadEmployees() {
		setLoading(true);
		try {
			const employees = await sharePointService.searchPeople('*');
			debugger

			setEmployees(employees);
			setLoading(false);
		} catch (error) {
			console.error('Failed to load employees:', error);
			setError(error instanceof Error ? error.message : "Failed to load employees");
		}
	}
	
	const departments = useMemo(() => {
		setSearchTerm('*');
    const depts = new Set(employees.map((emp: any) => emp.department).filter(Boolean))
    return Array.from(depts).sort()
  }, [employees])

  const filteredEmployees = useMemo(() => {
    return employees.filter((employee: any) => {
      const matchesSearch =
        !searchTerm ||
        employee.title.toLowerCase().includes(searchTerm.toLowerCase()) ||
        employee.role.toLowerCase().includes(searchTerm.toLowerCase()) ||
        employee.email.toLowerCase().includes(searchTerm.toLowerCase())

      const matchesDepartment = selectedDepartment === "all" || employee.department === selectedDepartment

      return matchesSearch && matchesDepartment
    })
  }, [employees, searchTerm, selectedDepartment])

	useEffect(() => {
		loadEmployees();
	}, []);

  const commandBarItems: ICommandBarItemProps[] = [
    {
      key: "refresh",
      text: "Refresh",
      iconProps: { iconName: "Refresh" },
      onClick: () => {},
    },
    {
      key: "viewMode",
      text: "View",
      iconProps: { iconName: "View" },
      subMenuProps: {
        items: [
          {
            key: "cards",
            text: "Cards",
            iconProps: { iconName: "GridViewMedium" },
            onClick: () => setViewMode("cards"),
          },
          {
            key: "list",
            text: "List",
            iconProps: { iconName: "List" },
            onClick: () => setViewMode("list"),
          },
          {
            key: "people",
            text: "People",
            iconProps: { iconName: "People" },
            onClick: () => setViewMode("people"),
          },
        ],
      },
    },
  ]

  const farItems: ICommandBarItemProps[] = [
    {
      key: "info",
      text: `${filteredEmployees.length} employees`,
      iconProps: { iconName: "Info" },
    },
  ]

  const openEmployeePanel = (employee: SharePointUser) => {
    setSelectedEmployee(employee)
    setIsPanelOpen(true)
  }

  if (loading) {
    return (
      <Stack horizontalAlign="center" verticalAlign="center" style={{ height: "400px" }}>
        <Spinner size={SpinnerSize.large} label="Loading employees from SharePoint..." />
      </Stack>
    )
  }

  if (error) {
    return (
      <Stack tokens={{ childrenGap: 16 }}>
        <MessageBar messageBarType={MessageBarType.error} isMultiline>
          <Text>{error}</Text>
        </MessageBar>
        <PrimaryButton text="Retry" iconProps={{ iconName: "Refresh" }} />
      </Stack>
    )
  }

  return (
    <Stack tokens={{ childrenGap: 20 }}>
      {/* Header */}
      <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
        <Stack>
          <Text variant="xxLarge" style={{ fontWeight: 600 }}>
            Employee Directory
          </Text>
          <Text variant="medium" style={{ color: "#605e5c" }}>
            Powered by SharePoint
          </Text>
        </Stack>
      </Stack>

      {/* Search and Filters */}
      <Stack tokens={{ childrenGap: 16 }}>
        <SearchBox
          placeholder="Search employees..."
          value={searchTerm}
          showIcon
        />

        <Pivot
          selectedKey={selectedDepartment}
          onLinkClick={(item) => setSelectedDepartment(item?.props.itemKey || "all")}
        >
          <PivotItem headerText="All Departments" itemKey="all" />
          {departments.map((dept) => (
            <PivotItem key={dept} headerText={dept} itemKey={dept} />
          ))}
        </Pivot>
      </Stack>

      {/* Command Bar */}
      <CommandBar items={commandBarItems} farItems={farItems} />

      {/* Content based on view mode */}
      {viewMode === "list" && <EmployeeListView employees={filteredEmployees} onEmployeeClick={openEmployeePanel} />}
      {viewMode === "cards" && <EmployeeCardsView employees={filteredEmployees} onEmployeeClick={openEmployeePanel} />}
      {viewMode === "people" && (
        <EmployeePeopleView employees={filteredEmployees} onEmployeeClick={openEmployeePanel} />
      )}

      {/* Employee Detail Panel */}
      <Panel
        isOpen={isPanelOpen}
        onDismiss={() => setIsPanelOpen(false)}
        type={PanelType.medium}
        headerText="Employee Details"
        closeButtonAriaLabel="Close"
      >
        {selectedEmployee && <EmployeeDetailPanel employee={selectedEmployee} />}
      </Panel>

      {filteredEmployees.length === 0 && <EmptyState />}
    </Stack>
  )
}

// List View Component
interface EmployeeListViewProps {
  employees: SharePointUser[]
  onEmployeeClick: (employee: SharePointUser) => void
}

const EmployeeListView: React.FC<EmployeeListViewProps> = ({ employees, onEmployeeClick }) => {
  const columns: IColumn[] = [
    {
      key: "name",
      name: "Name",
      fieldName: "name",
      minWidth: 200,
      maxWidth: 300,
      isResizable: true,
      onRender: (item: SharePointUser) => (
        <Persona
          text={item.title}
          secondaryText={item.jobTitle}
          imageUrl={item.pictureUrl}
          size={PersonaSize.size32}
          presence={PersonaPresence.online}
        />
      ),
    },
    {
      key: "department",
      name: "Department",
      fieldName: "department",
      minWidth: 120,
      maxWidth: 180,
      isResizable: true,
    },
    {
      key: "email",
      name: "Email",
      fieldName: "email",
      minWidth: 200,
      maxWidth: 300,
      isResizable: true,
      onRender: (item: SharePointUser) => (
        <a href={`mailto:${item.email}`} style={{ color: "#0078d4" }}>
          {item.email}
        </a>
      ),
    },
    {
      key: "phone",
      name: "Phone",
      fieldName: "phone",
      minWidth: 120,
      maxWidth: 180,
      isResizable: true,
    },
    {
      key: "location",
      name: "Location",
      fieldName: "location",
      minWidth: 120,
      maxWidth: 180,
      isResizable: true,
    },
  ]

  return (
    <DetailsList
      items={employees}
      columns={columns}
      layoutMode={DetailsListLayoutMode.justified}
      selectionMode={SelectionMode.none}
      onItemInvoked={onEmployeeClick}
    />
  )
}

// Cards View Component
interface EmployeeCardsViewProps {
  employees: SharePointUser[]
  onEmployeeClick: (employee: SharePointUser) => void
}

const EmployeeCardsView: React.FC<EmployeeCardsViewProps> = ({ employees, onEmployeeClick }) => {
  return (
    <div
      style={{
        display: "grid",
        gridTemplateColumns: "repeat(auto-fill, minmax(300px, 1fr))",
        gap: "16px",
      }}
    >
      {employees.map((employee) => (
        <DocumentCard key={employee.id} onClick={() => onEmployeeClick(employee)} style={{ cursor: "pointer" }}>
          <DocumentCardPreview
            {...({
              previewImages: [
                {
                  previewImageSrc: employee.pictureUrl || "/placeholder.svg?height=150&width=300",
                  imageFit: ImageFit.cover,
                  width: 300,
                  height: 150,
                },
              ],
            } as IDocumentCardPreviewProps)}
          />
          <DocumentCardDetails>
            <DocumentCardTitle title={employee.title} shouldTruncate />
            <DocumentCardActivity
              activity={employee.jobTitle}
              people={[
                {
                  name: employee.title,
                  profileImageSrc: employee.pictureUrl,
                },
              ]}
            />
          </DocumentCardDetails>
        </DocumentCard>
      ))}
    </div>
  )
}

// People View Component
interface EmployeePeopleViewProps {
  employees: SharePointUser[]
  onEmployeeClick: (employee: SharePointUser) => void
}

const EmployeePeopleView: React.FC<EmployeePeopleViewProps> = ({ employees, onEmployeeClick }) => {
  return (
    <div
      style={{
        display: "grid",
        gridTemplateColumns: "repeat(auto-fill, minmax(250px, 1fr))",
        gap: "16px",
      }}
    >
      {employees.map((employee) => (
        <div
          key={employee.id}
          onClick={() => onEmployeeClick(employee)}
          style={{
            backgroundColor: "#ffffff",
            border: "1px solid #edebe9",
            borderRadius: "4px",
            padding: "16px",
            cursor: "pointer",
            boxShadow: "0 1px 3px rgba(0,0,0,0.1)",
            transition: "all 0.2s ease",
          }}
          onMouseEnter={(e) => {
            e.currentTarget.style.boxShadow = "0 4px 8px rgba(0,0,0,0.15)"
            e.currentTarget.style.transform = "translateY(-2px)"
          }}
          onMouseLeave={(e) => {
            e.currentTarget.style.boxShadow = "0 1px 3px rgba(0,0,0,0.1)"
            e.currentTarget.style.transform = "translateY(0)"
          }}
        >
          <Stack horizontalAlign="center" tokens={{ childrenGap: 12 }}>
            <Persona
              text={employee.title}
              secondaryText={employee.jobTitle}
              tertiaryText={employee.department}
              imageUrl={employee.pictureUrl}
              size={PersonaSize.size72}
              presence={PersonaPresence.online}
            />
            <Stack tokens={{ childrenGap: 4 }}>
              <Text variant="small" style={{ color: "#605e5c" }}>
                📧 {employee.email}
              </Text>
              {employee.workPhone && (
                <Text variant="small" style={{ color: "#605e5c" }}>
                  📞 {employee.workPhone}
                </Text>
              )}
              <Text variant="small" style={{ color: "#605e5c" }}>
                📍 {employee.office}
              </Text>
            </Stack>
          </Stack>
        </div>
      ))}
    </div>
  )
}

// Employee Detail Panel Component
interface EmployeeDetailPanelProps {
  employee: SharePointUser
}

const EmployeeDetailPanel: React.FC<EmployeeDetailPanelProps> = ({ employee }) => {
  return (
    <Stack tokens={{ childrenGap: 20 }}>
      <Stack horizontalAlign="center">
        <Persona
          text={employee.title}
          secondaryText={employee.jobTitle}
          tertiaryText={employee.department}
          imageUrl={employee.pictureUrl}
          size={PersonaSize.size100}
          presence={PersonaPresence.online}
        />
      </Stack>

      <Stack tokens={{ childrenGap: 16 }}>
        <Stack>
          <Text variant="mediumPlus" style={{ fontWeight: 600 }}>
            Contact Information
          </Text>
          <Stack tokens={{ childrenGap: 8 }} style={{ marginTop: 8 }}>
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <Text style={{ minWidth: 60, color: "#605e5c" }}>Email:</Text>
              <a href={`mailto:${employee.email}`} style={{ color: "#0078d4" }}>
                {employee.email}
              </a>
            </Stack>
            {employee.workPhone && (
              <Stack horizontal tokens={{ childrenGap: 8 }}>
                <Text style={{ minWidth: 60, color: "#605e5c" }}>Phone:</Text>
                <a href={`tel:${employee.workPhone}`} style={{ color: "#0078d4" }}>
                  {employee.workPhone}
                </a>
              </Stack>
            )}
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <Text style={{ minWidth: 60, color: "#605e5c" }}>Location:</Text>
              <Text>{employee.office}</Text>
            </Stack>
          </Stack>
        </Stack>

        <Stack>
          <Text variant="mediumPlus" style={{ fontWeight: 600 }}>
            Organization
          </Text>
          <Stack tokens={{ childrenGap: 8 }} style={{ marginTop: 8 }}>
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <Text style={{ minWidth: 80, color: "#605e5c" }}>Department:</Text>
              <Text>{employee.department}</Text>
            </Stack>
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <Text style={{ minWidth: 80, color: "#605e5c" }}>Role:</Text>
              <Text>{employee.jobTitle}</Text>
            </Stack>
    
          </Stack>
        </Stack>


      </Stack>

      <Stack horizontal tokens={{ childrenGap: 8 }}>
        <PrimaryButton text="Send Email" iconProps={{ iconName: "Mail" }} href={`mailto:${employee.email}`} />
        {employee.workPhone && (
          <DefaultButton text="Call" iconProps={{ iconName: "Phone" }} href={`tel:${employee.workPhone}`} />
        )}
      </Stack>
    </Stack>
  )
}

// Empty State Component
const EmptyState: React.FC = () => {
  return (
    <Stack horizontalAlign="center" tokens={{ childrenGap: 16 }} style={{ padding: "40px" }}>
      <div style={{ fontSize: "48px" }}>👥</div>
      <Text variant="large" style={{ fontWeight: 600 }}>
        No employees found
      </Text>
      <Text style={{ color: "#605e5c", textAlign: "center" }}>
        Try adjusting your search or filter criteria to find employees.
      </Text>
    </Stack>
  )
}

