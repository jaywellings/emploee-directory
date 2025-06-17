import type { SharePointUser, SharePointUserProfile } from "../types/types"
import { spfi } from "@pnp/sp"
import { LogLevel, PnPLogging } from "@pnp/logging"
import "@pnp/sp/webs"
import "@pnp/sp/lists"
import "@pnp/sp/items"
import "@pnp/sp/site-users/web"
import "@pnp/sp/profiles"
import "@pnp/sp/search"
import { ISiteUserInfo } from "@pnp/sp/site-users/types"

export const sp = spfi("https://weldea.sharepoint.com").using(PnPLogging(LogLevel.Warning))

export class SharePointService {
  /**
   * Get all site users
   */
  async getSiteUsers(): Promise<SharePointUser[]> {
    try {
      const users = await sp.web.siteUsers
        .select("Id", "Title", "Email", "LoginName", "PrincipalType", "IsSiteAdmin")
        .filter("PrincipalType eq 1")() // Only get users, not groups

      return users.map((user:ISiteUserInfo) => ({
        id: user.Id.toString(),
        title: user.Title,
        email: user.Email,
        loginName: user.LoginName,
        principalType: user.PrincipalType,
				isSiteAdmin: user.IsSiteAdmin,
				jobTitle: '',
				pictureUrl: ''
      }))
    } catch (error) {
      console.error("Error fetching site users:", error)
      throw new Error("Failed to fetch site users from SharePoint")
    }
  }

  /**
   * Get user profile information
   */
  async getUserProfile(loginName: string): Promise<SharePointUserProfile | null> {
    try {
      const profile = await sp.profiles.getPropertiesFor(loginName)

      const getProperty = (key: string): string => {
        const prop = profile.UserProfileProperties?.find((p:any) => p.Key === key)
        return prop?.Value || ""
      }

      return {
        accountName: profile.AccountName,
        displayName: profile.DisplayName,
        email: profile.Email,
        personalUrl: profile.PersonalUrl,
        pictureUrl: profile.PictureUrl,
        title: getProperty("Title"),
        department: getProperty("Department"),
        office: getProperty("Office"),
        workPhone: getProperty("WorkPhone"),
        cellPhone: getProperty("CellPhone"),
        aboutMe: getProperty("AboutMe"),
        skills: getProperty("SPS-Skills"),
        manager: getProperty("Manager"),
      }
    } catch (error) {
      console.error(`Error fetching user profile for ${loginName}:`, error)
      return null
    }
  }

  /**
   * Search for people using SharePoint search
   */
  async searchPeople(query: string): Promise<SharePointUser[]> {
    try {
      const searchResults = await sp.search({
        Querytext: `${query}*`,
        SourceId: "b09a7990-05ea-4af9-81ef-edfab16c4e31", // People search source
        RowLimit: 50,
        SelectProperties: [
          "Title",
          "PreferredName",
          "WorkEmail",
          "PictureURL",
          "Department",
          "JobTitle",
          "OfficeNumber",
          "WorkPhone",
          "Path",
        ],
      })

      return searchResults.PrimarySearchResults.map((result: any) => ({
        id: result.Path.replace(/.*\//, ""),
        title: result.PreferredName || result.Title,
        email: result.WorkEmail,
        loginName: result.Path,
        department: result.Department,
        jobTitle: result.JobTitle,
        office: result.OfficeNumber,
        workPhone: result.WorkPhone,
        pictureUrl: result.PictureURL,
        principalType: 1,
        isSiteAdmin: false,
      }))
    } catch (error) {
      console.error("Error searching people:", error)
      throw new Error("Failed to search people in SharePoint")
    }
  }

  /**
   * Get people from a specific SharePoint list (if you have a custom employee list)
   */
  async getPeopleFromList(listName = "Employees"): Promise<any[]> {
    try {
      const items = await sp.web.lists
        .getByTitle(listName)
        .items.select(
          "Id",
          "Title",
          "Email",
          "Department",
          "JobTitle",
          "Phone",
          "Office",
          "Manager/Title",
          "StartDate",
          "ProfilePicture",
        )
        .expand("Manager")()

      return items.map((item) => ({
        id: item.Id.toString(),
        name: item.Title,
        email: item.Email,
        department: item.Department,
        role: item.JobTitle,
        phone: item.Phone,
        office: item.Office,
        manager: item.Manager?.Title,
        startDate: item.StartDate,
        avatar: item.ProfilePicture?.Url,
      }))
    } catch (error) {
      console.error(`Error fetching from list ${listName}:`, error)
      throw new Error(`Failed to fetch data from SharePoint list: ${listName}`)
    }
  }
}

export const sharePointService = new SharePointService()
