"use client"

import { useState, useEffect, useCallback } from "react"
import { SharePointUser } from "../types/types"
import { sharePointService } from "../service/dataService"

export function useSharePointUsers() {
  const [users, setUsers] = useState<SharePointUser[]>([])
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState<string | null>(null)

  const fetchUsers = useCallback(async () => {
    try {
      setLoading(true)
      const sharePointUsers = await sharePointService.getSiteUsers()
      setUsers(sharePointUsers)
      setError(null)
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to fetch users")
      console.error("Error fetching SharePoint users:", err)
    } finally {
      setLoading(false)
    }
  }, [])

  useEffect(() => {
    fetchUsers()
  }, [fetchUsers])

  return { users, loading, error, refetch: fetchUsers }
}



export function useSharePointSearch() {
  const [employees, setEmployees] = useState<SharePointUser[]>([])
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [isSearching, setIsSearching] = useState(false)

  const searchPeople = useCallback(async (query: string) => {


    try {
      setLoading(true)
      setIsSearching(true)
			const results = await sharePointService.searchPeople('*')
			console.log(results)
      setEmployees(results)
      setError(null)
      return results
    } catch (err) {
      const errorMessage = err instanceof Error ? err.message : "Search failed"
      setError(errorMessage)
      console.error("Error searching people:", err)
      return []
    } finally {
      setLoading(false)
    }
  }, [])

  const clearSearch = useCallback(() => {
    setEmployees([])
    setIsSearching(false)
    setError(null)
  }, [])

  return {
    employees,
    loading,
    error,
    isSearching,
    searchPeople,
    clearSearch,
  }
}
