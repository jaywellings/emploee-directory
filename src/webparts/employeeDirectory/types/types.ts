export interface Employee {
  id: string
  name: string
  role: string
  email: string
  phone?: string
  avatar?: string
  department: string
  startDate: string
  location: string
}

export interface Team {
  name: string
  description: string
  lead: string
  members: Employee[]
}

export interface SharePointUserProfile {
  accountName: string
  displayName: string
  email: string
  personalUrl: string
  pictureUrl: string
  title: string
  department: string
  office: string
  workPhone: string
  cellPhone: string
  aboutMe: string
  skills: string
  manager: string
}


export interface SharePointUser {
  id: string
  title: string
  email: string
  loginName: string
  principalType: number
  isSiteAdmin: boolean
  department?: string
  jobTitle: string
  office?: string
  workPhone?: string
  pictureUrl: string
}

export interface SharePointUserProfile {
  accountName: string
  displayName: string
  email: string
  personalUrl: string
  pictureUrl: string
  title: string
  department: string
  office: string
  workPhone: string
  cellPhone: string
  aboutMe: string
  skills: string
  manager: string
}

export interface SharePointEmployee {
  id: string
  name: string
  role: string
  email: string
  phone?: string
  avatar: string
  department: string
  startDate?: string
  location: string
  manager?: string
  skills?: string[]
}
