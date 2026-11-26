import { useEffect, useMemo, useState } from 'react'
import {
  Activity,
  Award,
  ChevronDown,
  ChevronRight,
  CircleHelp,
  ClipboardList,
  Download,
  Dumbbell,
  Eye,
  LogIn,
  LogOut,
  Medal,
  Menu,
  Plus,
  Search,
  ShieldCheck,
  Trash2,
  Trophy,
  Users,
  X,
} from 'lucide-react'
import { isSupabaseConfigured, supabase } from './lib/supabase'
import './App.css'

const sports = [
  { name: 'Dashboard', icon: Activity },
  { name: 'Teams', icon: Users },
  { name: 'Badminton', icon: Dumbbell },
  { name: 'Table Tennis', icon: Activity },
  { name: 'Carrom', icon: CircleHelp },
  { name: 'Chess', icon: ClipboardList },
  { name: 'Volleyball', icon: Trophy },
  { name: 'Football', icon: Trophy },
  { name: 'Cricket', icon: Trophy },
  { name: 'Points Table', icon: Award },
]
const bracketTeamColors = [
  { accent: '#247ba0', tint: '#eaf5f8' },
  { accent: '#b87718', tint: '#fbf2e3' },
  { accent: '#4c8a58', tint: '#edf6eb' },
  { accent: '#b6534c', tint: '#f9eeec' },
  { accent: '#755e9b', tint: '#f2eff8' },
  { accent: '#1f897f', tint: '#e9f5f3' },
]

function isByeFixture(match) {
  return match.match_status === 'bye'
    || (match.match_status === 'completed' && !match.participant2_id && match.winner_participant_id === match.participant1_id)
}

function createBracketSheet(XLSX, fixtureMatches, sportName, categoryName, teams) {
  const finalRoundNumber = Math.max(...fixtureMatches.map((match) => match.round_number))
  const maxRoundIndex = finalRoundNumber - 1
  const columnCount = finalRoundNumber * 3 - 1
  const bracketRows = []
  const setBracketCell = (rowIndex, columnIndex, value) => {
    bracketRows[rowIndex] ??= []
    bracketRows[rowIndex][columnIndex] = value
  }
  const getMatchRow = (roundIndex, matchIndex) => 2 ** (roundIndex + 1) + matchIndex * 4 * (2 ** roundIndex)
  const getMatchColumn = (roundIndex) => 1 + roundIndex * 3
  const border = { style: 'thin', color: { rgb: '8F9DA3' } }
  const connectorBorders = []
  const ensureCell = (sheet, rowIndex, columnIndex) => {
    const address = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex })
    sheet[address] ??= { t: 's', v: '' }
    return sheet[address]
  }
  const setCellBorder = (sheet, rowIndex, columnIndex, side) => {
    const cell = ensureCell(sheet, rowIndex, columnIndex)
    cell.s ??= {}
    cell.s.border ??= {}
    cell.s.border[side] = border
  }
  const getRoundName = (roundNumber) => roundNumber === finalRoundNumber ? 'Final'
    : roundNumber === finalRoundNumber - 1 ? 'Semi-finals'
      : roundNumber === finalRoundNumber - 2 ? 'Quarter-finals'
        : `Round ${roundNumber}`

  setBracketCell(0, 0, `${sportName} - ${categoryName}`)
  for (let roundNumber = 1; roundNumber <= finalRoundNumber; roundNumber += 1) {
    const roundIndex = roundNumber - 1
    const columnIndex = getMatchColumn(roundIndex)
      setBracketCell(1, columnIndex, getRoundName(roundNumber).toUpperCase())
    fixtureMatches.filter((match) => match.round_number === roundNumber).forEach((match, matchIndex) => {
      const rowIndex = getMatchRow(roundIndex, matchIndex)
      setBracketCell(rowIndex + 2, columnIndex, `Match ${match.match_number}`)
      ;[match.participant1, match.participant2].forEach((participant, slotIndex) => {
        const entry = participant
          ? `${participant.participant_name} - ${participant.teams?.team_name ?? 'Team'}`
          : isByeFixture(match) ? 'BYE' : 'Waiting for opponent'
        setBracketCell(rowIndex + slotIndex, columnIndex, entry)
      })
    })
  }
  for (let roundIndex = 0; roundIndex < maxRoundIndex; roundIndex += 1) {
    const currentRoundMatches = fixtureMatches.filter((match) => match.round_number === roundIndex + 1)
    const nextRoundMatches = fixtureMatches.filter((match) => match.round_number === roundIndex + 2)
    nextRoundMatches.forEach((_, matchIndex) => {
      if (!currentRoundMatches[matchIndex * 2 + 1]) return
      const firstChildCenter = getMatchRow(roundIndex, matchIndex * 2) + 1
      const secondChildCenter = getMatchRow(roundIndex, matchIndex * 2 + 1) + 1
      const nextMatchCenter = getMatchRow(roundIndex + 1, matchIndex) + 1
      const horizontalColumn = getMatchColumn(roundIndex) + 1
      const verticalColumn = getMatchColumn(roundIndex) + 2
      for (let rowIndex = firstChildCenter; rowIndex <= secondChildCenter; rowIndex += 1) {
        connectorBorders.push({ rowIndex, columnIndex: verticalColumn, side: 'right' })
      }
      ;[firstChildCenter, secondChildCenter].forEach((rowIndex) => {
        connectorBorders.push({ rowIndex, columnIndex: horizontalColumn, side: 'top' })
        connectorBorders.push({ rowIndex, columnIndex: verticalColumn, side: 'top' })
      })
      connectorBorders.push({ rowIndex: nextMatchCenter, columnIndex: verticalColumn, side: 'top' })
    })
  }

  const bracketSheet = XLSX.utils.aoa_to_sheet(bracketRows)
  connectorBorders.forEach(({ rowIndex, columnIndex, side }) => setCellBorder(bracketSheet, rowIndex, columnIndex, side))
  bracketSheet['!merges'] = [XLSX.utils.decode_range(`A1:${XLSX.utils.encode_col(columnCount - 1)}1`)]
  bracketSheet['!cols'] = Array.from({ length: columnCount }, (_, columnIndex) => ({
    wch: columnIndex === 0 ? 5 : (columnIndex - 1) % 3 === 0 ? 42 : 4,
  }))
  const matchRows = new Set()
  for (let roundIndex = 0; roundIndex <= maxRoundIndex; roundIndex += 1) {
    const roundMatches = fixtureMatches.filter((match) => match.round_number === roundIndex + 1)
    const columnIndex = getMatchColumn(roundIndex)
    roundMatches.forEach((match, matchIndex) => {
      const rowIndex = getMatchRow(roundIndex, matchIndex)
      matchRows.add(rowIndex)
      matchRows.add(rowIndex + 1)
      matchRows.add(rowIndex + 2)
      const matchLabelCell = ensureCell(bracketSheet, rowIndex + 2, columnIndex)
      matchLabelCell.s = { alignment: { horizontal: 'left', vertical: 'center', indent: 1 }, font: { name: 'Aptos', sz: 10, color: { rgb: '53616A' } } }
      ;[match.participant1, match.participant2].forEach((participant, slotIndex) => {
        const cell = ensureCell(bracketSheet, rowIndex + slotIndex, columnIndex)
        const teamColorIndex = teams.findIndex((team) => team.id === participant?.team_id)
        const teamColor = teamColorIndex < 0 ? null : bracketTeamColors[teamColorIndex % bracketTeamColors.length]
        const isWinner = participant?.id === match.winner_participant_id
        cell.s = {
          alignment: { horizontal: 'left', vertical: 'center', wrapText: true, indent: 1 },
          font: { name: 'Aptos', sz: 10, color: { rgb: '20313C' }, bold: isWinner },
          fill: teamColor ? { patternType: 'solid', fgColor: { rgb: teamColor.tint.slice(1) } } : { patternType: 'solid', fgColor: { rgb: 'FFFFFF' } },
          border: { top: border, bottom: border, left: border, right: border },
        }
      })
    })
  }
  const titleCell = bracketSheet.A1
  titleCell.s = { font: { name: 'Aptos Display', sz: 16, bold: true, color: { rgb: '17324D' } }, alignment: { vertical: 'center' } }
  for (let roundIndex = 0; roundIndex <= maxRoundIndex; roundIndex += 1) {
    const headerCell = ensureCell(bracketSheet, 1, getMatchColumn(roundIndex))
    headerCell.s = {
      font: { name: 'Aptos', sz: 9, bold: true, color: { rgb: '53616A' } },
      fill: { patternType: 'solid', fgColor: { rgb: 'EDF3F5' } },
      alignment: { horizontal: 'center', vertical: 'center' },
      border: { bottom: border },
    }
  }
  const lastBracketRow = Math.max(...matchRows)
  bracketSheet['!rows'] = Array.from({ length: lastBracketRow + 1 }, (_, rowIndex) => ({ hpt: matchRows.has(rowIndex) ? 20 : rowIndex < 2 ? 23 : 8 }))
  return bracketSheet
}

function App() {
  const [activePage, setActivePage] = useState(() => {
    const savedPage = localStorage.getItem('sportscarnival.activePage')
    return sports.some(({ name }) => name === savedPage) ? savedPage : 'Dashboard'
  })
  const [teams, setTeams] = useState([])
  const [players, setPlayers] = useState([])
  const [points, setPoints] = useState([])
  const [allCategories, setAllCategories] = useState([])
  const [pairCategories, setPairCategories] = useState([])
  const [pairRows, setPairRows] = useState([])
  const [fixtureMatches, setFixtureMatches] = useState([])
  const [tournamentResult, setTournamentResult] = useState(null)
  const [isEditingResult, setIsEditingResult] = useState(false)
  const [categoryId, setCategoryId] = useState('')
  const [pairTeamId, setPairTeamId] = useState('')
  const [user, setUser] = useState(null)
  const [searchText, setSearchText] = useState('')
  const [dialog, setDialog] = useState(null)
  const [notice, setNotice] = useState('')
  const [loading, setLoading] = useState(isSupabaseConfigured)
  const [isExportingSchedules, setIsExportingSchedules] = useState(false)
  const [sidebarOpen, setSidebarOpen] = useState(false)
  const [filterSport, setFilterSport] = useState('Overall')
  const isAdmin = user?.app_metadata?.role === 'admin'
  const activePairCategory = pairCategories.find((category) => category.id === categoryId)
  const pairRequiredGender = activePairCategory?.category_name.toLowerCase().startsWith("men's")
    ? 'Male'
    : activePairCategory?.category_name.toLowerCase().startsWith("women's") ? 'Female' : null
  const availablePairPlayers = players.filter((player) => player.team_id === pairTeamId && (!pairRequiredGender || player.gender === pairRequiredGender))

  useEffect(() => {
    localStorage.setItem('sportscarnival.activePage', activePage)
  }, [activePage])

  useEffect(() => {
    if (!isSupabaseConfigured) {
      return undefined
    }

    supabase.auth.getSession().then(({ data }) => setUser(data.session?.user ?? null))
    const { data: listener } = supabase.auth.onAuthStateChange((_event, session) => {
      setUser(session?.user ?? null)
    })
    return () => listener.subscription.unsubscribe()
  }, [])

  useEffect(() => {
    let active = true
    async function loadPublicData() {
      setLoading(true)
      if (!isSupabaseConfigured) {
        setTeams([])
        setPlayers([])
        setPoints([])
        setAllCategories([])
        setLoading(false)
        return
      }

      const [teamResult, playerResult, pointsResult, categoryResult] = await Promise.all([
        supabase.from('teams').select('id, team_name, department, logo_url, created_at').order('team_name'),
        supabase.from('players').select('id, team_id, player_name, gender').order('player_name'),
        supabase.from('points_table').select('team_id, sport_id, category_id, winner_points, runner_up_points, total_points, sports(sport_name), tournament_categories(category_name), teams(team_name)'),
        supabase.from('tournament_categories').select('id, sport_id, category_name, competition_type, sports!inner(sport_name)').order('category_name'),
      ])
      if (!active) return
      const failed = [teamResult, playerResult, pointsResult, categoryResult].find((result) => result.error)
      if (failed) setNotice(failed.error.message)
      setTeams(teamResult.data ?? [])
      setPlayers(playerResult.data ?? [])
      setPoints(pointsResult.data ?? [])
      setAllCategories(categoryResult.data ?? [])
      setLoading(false)
    }

    loadPublicData()
    return () => { active = false }
  }, [isAdmin])

  useEffect(() => {
    let active = true
    async function loadCategories() {
      if (!isSupabaseConfigured || !['Badminton', 'Carrom', 'Table Tennis', 'Chess', 'Volleyball', 'Football', 'Cricket'].includes(activePage)) {
        setPairCategories([])
        setPairRows([])
        setTournamentResult(null)
        setCategoryId('')
        return
      }
      const { data, error } = await supabase
        .from('tournament_categories')
        .select('id, sport_id, category_name, competition_type, sports!inner(id, sport_name)')
        .eq('sports.sport_name', activePage)
        .order('category_name')
      if (!active) return
      if (error) setNotice(error.message)
      const categories = data ?? []
      setPairCategories(categories)
      setCategoryId((current) => categories.some((category) => category.id === current) ? current : categories[0]?.id ?? '')
    }
    loadCategories()
    return () => { active = false }
  }, [activePage, isAdmin])

  useEffect(() => {
    let active = true
    async function loadPairs() {
      if (!isSupabaseConfigured || !categoryId) {
        setPairRows([])
        setFixtureMatches([])
        return
      }
      const [pairResult, fixtureResult, resultResult] = await Promise.all([
        supabase.from('tournament_participants')
          .select('id, category_id, team_id, participant_name, teams(team_name), participant_players(player_id, players(id, player_name, gender))')
          .eq('category_id', categoryId)
          .order('participant_name'),
        supabase.from('matches')
          .select('id, category_id, round_number, match_number, participant1_id, participant2_id, winner_participant_id, match_status, next_match_id, next_match_slot, participant1:tournament_participants!matches_participant1_id_fkey(id, participant_name, team_id, teams(team_name)), participant2:tournament_participants!matches_participant2_id_fkey(id, participant_name, team_id, teams(team_name))')
          .eq('category_id', categoryId)
          .order('round_number')
          .order('match_number'),
        supabase.from('tournament_results')
          .select('category_id, winner_team_id, runner_up_team_id, player_of_tournament_id, best_batsman_id, best_bowler_id')
          .eq('category_id', categoryId)
          .maybeSingle(),
      ])
      if (!active) return
      if (pairResult.error) setNotice(pairResult.error.message)
      if (fixtureResult.error) setNotice(fixtureResult.error.message)
      if (resultResult.error) setNotice(resultResult.error.message)
      setPairRows(pairResult.data ?? [])
      setFixtureMatches(fixtureResult.data ?? [])
      setTournamentResult(resultResult.data ?? null)
      setIsEditingResult(!resultResult.data)
    }
    loadPairs()
    return () => { active = false }
  }, [categoryId, isAdmin])

  const filteredTeams = useMemo(() => {
    const query = searchText.trim().toLowerCase()
    return teams.filter((team) => team.team_name.toLowerCase().includes(query)
      || players.some((player) => player.team_id === team.id
        && `${player.player_name} ${player.gender}`.toLowerCase().includes(query)))
  }, [teams, players, searchText])
  const totalPoints = useMemo(() => {
    const totals = new Map(teams.map((team) => [team.id, {
      team_name: team.team_name,
      winner_points: 0,
      runner_up_points: 0,
      total_points: 0,
    }]))
    points.forEach((row) => {
      const teamName = row.teams?.team_name ?? 'Team'
      const teamKey = row.team_id ?? teamName
      const current = totals.get(teamKey) ?? { team_name: teamName, winner_points: 0, runner_up_points: 0, total_points: 0 }
      current.winner_points += row.winner_points ?? 0
      current.runner_up_points += row.runner_up_points ?? 0
      current.total_points += row.total_points ?? 0
      totals.set(teamKey, current)
    })
    return [...totals.values()].sort((a, b) => b.total_points - a.total_points || a.team_name.localeCompare(b.team_name))
  }, [points, teams])

  async function saveRecord(event) {
    event.preventDefault()
    if (!isAdmin || !dialog) return
    const formData = new FormData(event.currentTarget)
    const values = Object.fromEntries(formData.entries())
    if (dialog.kind === 'pair') {
      await savePair(values)
      return
    }
    if (dialog.kind === 'participant') {
      await saveSingleParticipant(values)
      return
    }
    let result

    if (dialog.kind === 'team') {
      const payload = { team_name: values.team_name.trim() }
      result = dialog.record
        ? await supabase.from('teams').update(payload).eq('id', dialog.record.id)
        : await supabase.from('teams').insert(payload)
    } else {
      const payload = { player_name: values.player_name.trim(), gender: values.gender, team_id: values.team_id }
      result = dialog.record
        ? await supabase.from('players').update(payload).eq('id', dialog.record.id)
        : await supabase.from('players').insert(payload)
    }

    if (result.error) {
      setNotice(result.error.message)
      return
    }
    setDialog(null)
    setNotice('Changes saved.')
    refreshRoster()
  }

  async function savePair(values) {
    const category = pairCategories.find((item) => item.id === (dialog.record?.category_id ?? categoryId))
    const firstPlayer = players.find((player) => player.id === values.player1_id)
    const secondPlayer = players.find((player) => player.id === values.player2_id)
    if (!category || !firstPlayer || !secondPlayer) {
      setNotice('Choose a category and two registered players.')
      return
    }
    if (firstPlayer.id === secondPlayer.id) {
      setNotice('A player cannot be selected twice in one pair.')
      return
    }
    if (firstPlayer.team_id !== values.team_id || secondPlayer.team_id !== values.team_id) {
      setNotice('Both players must belong to the selected team.')
      return
    }
    const requiredGender = category.category_name.toLowerCase().startsWith("men's")
      ? 'Male'
      : category.category_name.toLowerCase().startsWith("women's") ? 'Female' : null
    if (requiredGender && (firstPlayer.gender !== requiredGender || secondPlayer.gender !== requiredGender)) {
      setNotice(`Select two ${requiredGender.toLowerCase()} players for this category.`)
      return
    }
    const selectedPlayerIds = [firstPlayer.id, secondPlayer.id].sort()
    const duplicate = pairRows.some((pair) => {
      if (pair.id === dialog.record?.id) return false
      const existingIds = pair.participant_players.map((item) => item.player_id).sort()
      return existingIds.length === 2 && existingIds[0] === selectedPlayerIds[0] && existingIds[1] === selectedPlayerIds[1]
    })
    if (duplicate) {
      setNotice('That player pair is already registered in this category.')
      return
    }

    if (fixtureMatches.length) {
      const confirmation = `Saving this pair will clear the existing fixtures, final result, and awarded points for ${category.category_name}. Continue?`
      if (!window.confirm(confirmation) || !(await resetFixturesForRosterChange(category.id))) return
    }

    const participantName = values.participant_name.trim() || `${firstPlayer.player_name} / ${secondPlayer.player_name}`
    const participantPayload = { category_id: category.id, team_id: values.team_id, participant_name: participantName }
    const participantResult = dialog.record
      ? await supabase.from('tournament_participants').update(participantPayload).eq('id', dialog.record.id).select('id').single()
      : await supabase.from('tournament_participants').insert(participantPayload).select('id').single()
    if (participantResult.error) {
      setNotice(participantResult.error.message)
      return
    }

    const participantId = participantResult.data.id
    const previousPlayerIds = dialog.record?.participant_players.map((item) => item.player_id) ?? []
    if (dialog.record) {
      const { error } = await supabase.from('participant_players').delete().eq('participant_id', participantId)
      if (error) {
        setNotice(error.message)
        return
      }
    }
    const { error: linkError } = await supabase.from('participant_players').insert([
      { participant_id: participantId, player_id: firstPlayer.id },
      { participant_id: participantId, player_id: secondPlayer.id },
    ])
    if (linkError) {
      if (dialog.record) {
        await supabase.from('participant_players').insert(previousPlayerIds.map((player_id) => ({ participant_id: participantId, player_id })))
      } else {
        await supabase.from('tournament_participants').delete().eq('id', participantId)
      }
      setNotice(linkError.message)
      return
    }
    setDialog(null)
    setNotice(dialog.record ? 'Pair updated.' : 'Pair added.')
    const { data } = await supabase
      .from('tournament_participants')
      .select('id, category_id, team_id, participant_name, teams(team_name), participant_players(player_id, players(id, player_name, gender))')
      .eq('category_id', category.id)
      .order('participant_name')
    setPairRows(data ?? [])
  }

  async function resetFixturesForRosterChange(categoryId) {
    const { error } = await supabase.rpc('reset_tournament_fixtures', { p_category_id: categoryId })
    if (error) {
      setNotice(error.message)
      return false
    }
    setFixtureMatches([])
    setTournamentResult(null)
    setPoints((current) => current.filter((row) => row.category_id !== categoryId))
    return true
  }

  async function deleteFixtures() {
    if (!isAdmin || !categoryId || !fixtureMatches.length) return
    const categoryName = activePairCategory?.category_name ?? 'this category'
    if (!window.confirm(`Delete the bracket for ${categoryName}? This also clears its saved result and awarded points. Registered players and pairs will remain.`)) return
    if (await resetFixturesForRosterChange(categoryId)) setNotice('Bracket deleted. Registered players and pairs were kept.')
  }

  async function saveSingleParticipant(values) {
    if (fixtureMatches.length) {
      const changeDescription = dialog.record ? 'Updating this player' : 'Adding this player'
      const categoryName = activePairCategory?.category_name ?? 'this category'
      if (!window.confirm(`${changeDescription} will clear the bracket, result, and awarded points for ${categoryName}. Continue?`)) return
      if (!(await resetFixturesForRosterChange(categoryId))) return
    }
    const category = pairCategories.find((item) => item.id === categoryId)
    const player = players.find((item) => item.id === values.player_id)
    if (!category || !player || player.team_id !== values.team_id) {
      setNotice('Choose a registered player from the selected team.')
      return
    }
    const requiredGender = category.category_name.toLowerCase().startsWith("men's")
      ? 'Male'
      : category.category_name.toLowerCase().startsWith("women's") ? 'Female' : null
    if (requiredGender && player.gender !== requiredGender) {
      setNotice(`Select a ${requiredGender.toLowerCase()} player for this category.`)
      return
    }
    if (pairRows.some((participant) => participant.id !== dialog.record?.id
      && participant.participant_players.some((item) => item.player_id === player.id))) {
      setNotice('This player is already registered in this category.')
      return
    }
    const payload = { category_id: category.id, team_id: player.team_id, participant_name: player.player_name }
    const participantResult = dialog.record
      ? await supabase.from('tournament_participants').update(payload).eq('id', dialog.record.id).select('id').single()
      : await supabase.from('tournament_participants').insert(payload).select('id').single()
    if (participantResult.error) {
      setNotice(participantResult.error.message)
      return
    }
    const previousPlayerIds = dialog.record?.participant_players.map((item) => item.player_id) ?? []
    if (dialog.record) {
      const { error: deleteError } = await supabase.from('participant_players').delete().eq('participant_id', dialog.record.id)
      if (deleteError) {
        setNotice(deleteError.message)
        return
      }
    }
    const { error } = await supabase.from('participant_players').insert({ participant_id: participantResult.data.id, player_id: player.id })
    if (error) {
      if (dialog.record) {
        await supabase.from('participant_players').insert(previousPlayerIds.map((player_id) => ({ participant_id: participantResult.data.id, player_id })))
      } else {
        await supabase.from('tournament_participants').delete().eq('id', participantResult.data.id)
      }
      setNotice(error.message)
      return
    }
    setDialog(null)
    setNotice(dialog.record ? 'Participant updated.' : 'Participant added.')
    const { data } = await supabase.from('tournament_participants')
      .select('id, category_id, team_id, participant_name, teams(team_name), participant_players(player_id, players(id, player_name, gender))')
      .eq('category_id', category.id)
      .order('participant_name')
    setPairRows(data ?? [])
  }

  async function deletePair(pair) {
    if (!isAdmin) return
    const { count, error: matchError } = await supabase
      .from('matches')
      .select('id', { count: 'exact', head: true })
      .eq('category_id', pair.category_id)
    if (matchError) {
      setNotice(matchError.message)
      return
    }
    const hasFixtures = count > 0
    const confirmation = hasFixtures
      ? `Delete ${pair.participant_name}? This will also clear the category's fixtures, final result, and awarded points.`
      : `Delete ${pair.participant_name}?`
    if (!window.confirm(confirmation)) return
    if (hasFixtures && !(await resetFixturesForRosterChange(pair.category_id))) return
    const { error } = await supabase.from('tournament_participants').delete().eq('id', pair.id)
    if (error) {
      setNotice(error.message)
      return
    }
    setNotice('Pair deleted.')
    setPairRows((current) => current.filter((item) => item.id !== pair.id))
  }

  function openPairDialog(record = null) {
    if (!isAdmin) return
    setPairTeamId(record?.team_id ?? '')
    setDialog({ kind: 'pair', record })
  }

  function openParticipantDialog(record = null) {
    if (!isAdmin) return
    setPairTeamId(record?.team_id ?? '')
    setDialog({ kind: 'participant', record })
  }

  async function generateFixtures() {
    if (!isAdmin || !categoryId || pairRows.length < 2) return
    const regenerating = fixtureMatches.length > 0
    const confirmation = regenerating
      ? `Regenerate this bracket? Existing matches, results, and points for ${activePairCategory?.category_name ?? 'this category'} will be cleared.`
      : `Generate knockout fixtures for ${pairRows.length} entries?`
    if (!window.confirm(confirmation)) return
    if (regenerating) {
      const { error } = await supabase.rpc('reset_tournament_fixtures', { p_category_id: categoryId })
      if (error) {
        setNotice(error.message)
        return
      }
      setFixtureMatches([])
      setTournamentResult(null)
      const { data: updatedPoints } = await supabase.from('points_table')
        .select('team_id, sport_id, category_id, winner_points, runner_up_points, total_points, sports(sport_name), tournament_categories(category_name), teams(team_name)')
      setPoints(updatedPoints ?? [])
    }

    const shuffledEntries = [...pairRows]
    const shuffle = (items) => {
      const result = [...items]
      for (let index = result.length - 1; index > 0; index -= 1) {
        const otherIndex = Math.floor(Math.random() * (index + 1))
        ;[result[index], result[otherIndex]] = [result[otherIndex], result[index]]
      }
      return result
    }
    const takeLargestTeamEntry = (entries, excludeTeamId = null) => {
      const counts = new Map()
      entries.forEach((entry) => counts.set(entry.team_id, (counts.get(entry.team_id) ?? 0) + 1))
      const eligibleTeams = [...counts.entries()]
        .filter(([teamId]) => teamId !== excludeTeamId)
      const sourceTeams = eligibleTeams.length ? eligibleTeams : [...counts.entries()]
      const largestCount = Math.max(...sourceTeams.map(([, count]) => count))
      const largestTeams = sourceTeams.filter(([, count]) => count === largestCount).map(([teamId]) => teamId)
      const selectedTeamId = largestTeams[Math.floor(Math.random() * largestTeams.length)]
      const teamEntries = entries.filter((entry) => entry.team_id === selectedTeamId)
      const selectedEntry = teamEntries[Math.floor(Math.random() * teamEntries.length)]
      entries.splice(entries.findIndex((entry) => entry.id === selectedEntry.id), 1)
      return selectedEntry
    }

    const bracketSize = 2 ** Math.ceil(Math.log2(shuffledEntries.length))
    const roundCount = Math.log2(bracketSize)
    const bracketRounds = Array.from({ length: roundCount }, (_, roundIndex) => {
      const matchCount = bracketSize / (2 ** (roundIndex + 1))
      return Array.from({ length: matchCount }, (_, matchIndex) => ({
        id: crypto.randomUUID(),
        category_id: categoryId,
        round_number: roundIndex + 1,
        match_number: matchIndex + 1,
        participant1_id: null,
        participant2_id: null,
        winner_participant_id: null,
        match_status: 'upcoming',
        next_match_id: null,
        next_match_slot: null,
      }))
    })

    const teamSets = new Map()
    const byeCount = bracketSize - shuffledEntries.length
    const byeMatches = []
    for (let index = 0; index < byeCount; index += 1) {
      byeMatches.push([takeLargestTeamEntry(shuffledEntries), null])
    }
    const firstRoundMatches = []
    while (shuffledEntries.length) {
      const firstEntry = takeLargestTeamEntry(shuffledEntries)
      const opponent = takeLargestTeamEntry(shuffledEntries, firstEntry.team_id)
      firstRoundMatches.push([firstEntry, opponent])
    }
    const initialMatchups = shuffle([...byeMatches, ...firstRoundMatches])

    initialMatchups.forEach(([firstEntry, secondEntry], matchIndex) => {
      const match = bracketRounds[0][matchIndex]
      match.participant1_id = firstEntry?.id ?? null
      match.participant2_id = secondEntry?.id ?? null
      if (firstEntry && !secondEntry) {
        match.match_status = 'bye'
      }
      teamSets.set(match.id, new Set([firstEntry?.team_id, secondEntry?.team_id].filter(Boolean)))
    })

    const groupMatchesForNextRound = (round) => {
      const remaining = shuffle(round)
      const ordered = []
      while (remaining.length) {
        const firstMatch = remaining.shift()
        ordered.push(firstMatch)
        if (!remaining.length) break
        const firstTeams = teamSets.get(firstMatch.id) ?? new Set()
        const firstIsBye = firstMatch.match_status === 'bye'
        const hasDifferentByeStatus = remaining.some((candidate) => (candidate.match_status === 'bye') !== firstIsBye)
        let bestOverlap = Infinity
        let bestIndexes = []
        remaining.forEach((candidate, candidateIndex) => {
          if (hasDifferentByeStatus && (candidate.match_status === 'bye') === firstIsBye) return
          const candidateTeams = teamSets.get(candidate.id) ?? new Set()
          const overlap = [...firstTeams].filter((teamId) => candidateTeams.has(teamId)).length
          if (overlap < bestOverlap) {
            bestOverlap = overlap
            bestIndexes = [candidateIndex]
          } else if (overlap === bestOverlap) {
            bestIndexes.push(candidateIndex)
          }
        })
        const selectedIndex = bestIndexes[Math.floor(Math.random() * bestIndexes.length)]
        ordered.push(remaining.splice(selectedIndex, 1)[0])
      }
      ordered.forEach((match, index) => { match.match_number = index + 1 })
      return ordered
    }

    for (let roundIndex = 1; roundIndex < bracketRounds.length; roundIndex += 1) {
      const previousRound = groupMatchesForNextRound(bracketRounds[roundIndex - 1])
      bracketRounds[roundIndex - 1] = previousRound
      bracketRounds[roundIndex].forEach((match, matchIndex) => {
        const children = [previousRound[matchIndex * 2], previousRound[matchIndex * 2 + 1]]
        match.participant1_id = children[0]?.winner_participant_id ?? null
        match.participant2_id = children[1]?.winner_participant_id ?? null
        teamSets.set(match.id, new Set(children.flatMap((child) => [...(teamSets.get(child?.id) ?? [])])))
      })
    }
    for (let roundIndex = 0; roundIndex < bracketRounds.length - 1; roundIndex += 1) {
      bracketRounds[roundIndex].forEach((match, matchIndex) => {
        match.next_match_id = bracketRounds[roundIndex + 1][Math.floor(matchIndex / 2)].id
        match.next_match_slot = matchIndex % 2 + 1
      })
    }

    for (let roundIndex = bracketRounds.length - 1; roundIndex >= 0; roundIndex -= 1) {
      const { error } = await supabase.from('matches').insert(bracketRounds[roundIndex])
      if (error) {
        await supabase.from('matches').delete().eq('category_id', categoryId)
        setNotice(error.message)
        return
      }
    }
    const { data, error } = await supabase.from('matches')
      .select('id, category_id, round_number, match_number, participant1_id, participant2_id, winner_participant_id, match_status, next_match_id, next_match_slot, participant1:tournament_participants!matches_participant1_id_fkey(id, participant_name, team_id, teams(team_name)), participant2:tournament_participants!matches_participant2_id_fkey(id, participant_name, team_id, teams(team_name))')
      .eq('category_id', categoryId)
      .order('round_number')
      .order('match_number')
    if (error) {
      setNotice(error.message)
      return
    }
    setFixtureMatches(data ?? [])
    setNotice(regenerating ? 'Fixtures regenerated.' : 'Knockout fixtures generated. Same-team matchups are delayed where bracket seeding allows.')
  }

  async function setMatchWinner(match, winnerId) {
    if (!isAdmin || !winnerId || ![match.participant1_id, match.participant2_id].includes(winnerId)) return
    const nextMatch = fixtureMatches.find((item) => item.id === match.next_match_id)
    if (nextMatch?.match_status === 'completed') {
      setNotice('This result cannot be changed after the next match is completed.')
      return
    }
    const { error } = await supabase.from('matches').update({
      winner_participant_id: winnerId,
      match_status: 'completed',
    }).eq('id', match.id)
    if (error) {
      setNotice(error.message)
      return
    }

    if (nextMatch) {
      const slotColumn = match.next_match_slot === 1 ? 'participant1_id' : 'participant2_id'
      const { error: advanceError } = await supabase.from('matches').update({ [slotColumn]: winnerId }).eq('id', nextMatch.id)
      if (advanceError) {
        setNotice(advanceError.message)
        return
      }
    } else {
      const winnerParticipant = [match.participant1, match.participant2].find((participant) => participant?.id === winnerId)
      const runnerUpParticipant = [match.participant1, match.participant2].find((participant) => participant?.id !== winnerId)
      if (winnerParticipant && runnerUpParticipant) {
        const saved = await persistTournamentResult({
          category_id: categoryId,
          winner_team_id: winnerParticipant.team_id,
          runner_up_team_id: runnerUpParticipant.team_id,
        })
        if (!saved) return
      }
    }

    const { data } = await supabase.from('matches')
      .select('id, category_id, round_number, match_number, participant1_id, participant2_id, winner_participant_id, match_status, next_match_id, next_match_slot, participant1:tournament_participants!matches_participant1_id_fkey(id, participant_name, team_id, teams(team_name)), participant2:tournament_participants!matches_participant2_id_fkey(id, participant_name, team_id, teams(team_name))')
      .eq('category_id', categoryId)
      .order('round_number')
      .order('match_number')
    setFixtureMatches(data ?? [])
    setNotice('Match winner saved.')
  }

  async function setMatchIncomplete(match) {
    if (!isAdmin || match.match_status !== 'completed') return
    const { error } = await supabase.rpc('mark_match_incomplete', { p_match_id: match.id })
    if (error) {
      setNotice(error.message)
      return
    }
    const { data, error: refreshError } = await supabase.from('matches')
      .select('id, category_id, round_number, match_number, participant1_id, participant2_id, winner_participant_id, match_status, next_match_id, next_match_slot, participant1:tournament_participants!matches_participant1_id_fkey(id, participant_name, team_id, teams(team_name)), participant2:tournament_participants!matches_participant2_id_fkey(id, participant_name, team_id, teams(team_name))')
      .eq('category_id', categoryId)
      .order('round_number')
      .order('match_number')
    if (refreshError) {
      setNotice(refreshError.message)
      return
    }
    setFixtureMatches(data ?? [])
    if (!match.next_match_id) {
      setTournamentResult(null)
      setPoints((current) => current.filter((row) => row.category_id !== categoryId))
    }
    setNotice('Match marked incomplete.')
  }

  async function downloadAllSchedulesExcel() {
    if (!isSupabaseConfigured || isExportingSchedules) return
    setIsExportingSchedules(true)
    try {
      const { data: matches, error } = await supabase.from('matches')
        .select('id, category_id, round_number, match_number, participant1_id, participant2_id, winner_participant_id, match_status, participant1:tournament_participants!matches_participant1_id_fkey(id, participant_name, team_id, teams(team_name)), participant2:tournament_participants!matches_participant2_id_fkey(id, participant_name, team_id, teams(team_name)), tournament_categories!inner(category_name, sports!inner(sport_name))')
        .order('category_id')
        .order('round_number')
        .order('match_number')
      if (error) {
        setNotice(error.message)
        return
      }
      if (!matches?.length) {
        setNotice('No fixture schedules have been generated yet.')
        return
      }

      const XLSX = await import('xlsx-js-style')
      const matchesByCategory = new Map()
      matches.forEach((match) => {
        const sportName = match.tournament_categories?.sports?.sport_name ?? 'Sport'
        const categoryName = match.tournament_categories?.category_name ?? 'Fixtures'
        if (!matchesByCategory.has(match.category_id)) {
          matchesByCategory.set(match.category_id, { sportName, categoryName, matches: [] })
        }
        matchesByCategory.get(match.category_id).matches.push(match)
      })

      const workbook = XLSX.utils.book_new()
      const usedSheetNames = new Set()
      ;[...matchesByCategory.values()]
        .sort((first, second) => `${first.sportName} ${first.categoryName}`.localeCompare(`${second.sportName} ${second.categoryName}`))
        .forEach((group) => {
          const baseName = `${group.sportName === 'Table Tennis' ? 'TT' : group.sportName} - ${group.categoryName}`
            .replace(/[\\/?*:[\]]/g, '-')
            .slice(0, 31)
          let sheetName = baseName
          let suffix = 2
          while (usedSheetNames.has(sheetName.toLowerCase())) {
            const suffixText = ` (${suffix})`
            suffix += 1
            sheetName = `${baseName.slice(0, 31 - suffixText.length)}${suffixText}`
          }
          usedSheetNames.add(sheetName.toLowerCase())
          const sheet = createBracketSheet(XLSX, group.matches, group.sportName, group.categoryName, teams)
          XLSX.utils.book_append_sheet(workbook, sheet, sheetName)
      })
      XLSX.writeFile(workbook, 'Sports Carnival schedules.xlsx')
      setNotice('All available game schedules were exported to Excel.')
    } catch (error) {
      setNotice(error.message ?? 'Could not export the schedules.')
    } finally {
      setIsExportingSchedules(false)
    }
  }

  async function downloadFixturesExcel() {
    if (!fixtureMatches.length) return
    const XLSX = await import('xlsx-js-style')
    const finalRoundNumber = Math.max(...fixtureMatches.map((match) => match.round_number))
    const getRoundName = (roundNumber) => roundNumber === finalRoundNumber ? 'Final'
      : roundNumber === finalRoundNumber - 1 ? 'Semi-finals'
        : roundNumber === finalRoundNumber - 2 ? 'Quarter-finals'
          : `Round ${roundNumber}`
    const headers = ['Sport', 'Category', 'Round', 'Match', 'Status', 'Team 1', 'Entry 1', 'Team 2', 'Entry 2', 'Winner']
    const rows = fixtureMatches.map((match) => {
      const participants = [match.participant1, match.participant2]
      const winner = participants.find((participant) => participant?.id === match.winner_participant_id)
      const isBye = isByeFixture(match)
      return [
        activePage,
        activePairCategory?.category_name ?? '',
        getRoundName(match.round_number),
        match.match_number,
        isBye ? match.winner_participant_id ? 'Bye · Advanced' : 'Bye · Select to advance' : match.match_status === 'completed' ? 'Completed' : 'Upcoming',
        participants[0]?.teams?.team_name ?? '',
        participants[0]?.participant_name ?? (isBye ? 'BYE' : 'TBD'),
        participants[1]?.teams?.team_name ?? '',
        participants[1]?.participant_name ?? (isBye ? 'BYE' : 'TBD'),
        winner?.participant_name ?? '',
      ]
    })
    const detailSheet = XLSX.utils.aoa_to_sheet([headers, ...rows])
    detailSheet['!cols'] = headers.map((header, columnIndex) => ({
      wch: Math.min(36, Math.max(header.length + 2, ...rows.map((row) => String(row[columnIndex] ?? '').length + 2))),
    }))
    const bracketSheet = createBracketSheet(
      XLSX,
      fixtureMatches,
      activePage,
      activePairCategory?.category_name ?? 'Fixtures',
      teams,
    )
    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, bracketSheet, 'Bracket')
    XLSX.utils.book_append_sheet(workbook, detailSheet, 'Matches')
    const fileName = `${activePage} ${activePairCategory?.category_name ?? 'Fixtures'}`
      .replace(/[<>:"/\\|?*]/g, '-')
    XLSX.writeFile(workbook, `${fileName}.xlsx`)
  }

  async function persistTournamentResult(result) {
    const { error } = await supabase.rpc('save_tournament_result', {
      p_category_id: result.category_id,
      p_winner_team_id: result.winner_team_id,
      p_runner_up_team_id: result.runner_up_team_id,
      p_player_of_tournament_id: result.player_of_tournament_id ?? null,
      p_best_batsman_id: result.best_batsman_id ?? null,
      p_best_bowler_id: result.best_bowler_id ?? null,
    })
    if (error) {
      setNotice(error.message)
      return false
    }
    setTournamentResult(result)
    setIsEditingResult(false)
    const { data: updatedPoints } = await supabase.from('points_table')
      .select('team_id, sport_id, category_id, winner_points, runner_up_points, total_points, sports(sport_name), tournament_categories(category_name), teams(team_name)')
    setPoints(updatedPoints ?? [])
    setNotice('Winner, runner-up, and points saved.')
    return true
  }

  async function saveTeamSportResult(event) {
    event.preventDefault()
    if (!isAdmin || !categoryId) return
    const values = Object.fromEntries(new FormData(event.currentTarget).entries())
    if (!values.winner_team_id || !values.runner_up_team_id) {
      setNotice('Select both a winner and a runner-up.')
      return
    }
    if (values.winner_team_id === values.runner_up_team_id) {
      setNotice('Winner and runner-up must be different teams.')
      return
    }
    const saved = await persistTournamentResult({
      category_id: categoryId,
      winner_team_id: values.winner_team_id,
      runner_up_team_id: values.runner_up_team_id,
      player_of_tournament_id: values.player_of_tournament_id || null,
      best_batsman_id: values.best_batsman_id || null,
      best_bowler_id: values.best_bowler_id || null,
    })
    if (saved) setTournamentResult((current) => ({ ...current, category_id: categoryId }))
  }

  function renderTeamFinalStandings() {
    const winnerTeam = teams.find((team) => team.id === tournamentResult?.winner_team_id)
    const runnerUpTeam = teams.find((team) => team.id === tournamentResult?.runner_up_team_id)
    const awardPlayer = (playerId) => {
      const player = players.find((item) => item.id === playerId)
      const team = teams.find((item) => item.id === player?.team_id)
      return player ? `${player.player_name}${team ? ` (${team.team_name})` : ''}` : '—'
    }
    return <div className="final-standing-content">
      <div className="podium-grid">
        <article className="podium-card podium-gold"><span className="podium-medal"><Medal size={27} /></span><div className="podium-copy"><span>1ST PLACE</span><strong>Winner</strong><h3>{winnerTeam?.team_name ?? '—'}</h3></div><div className="podium-points"><strong>4</strong><span>points</span></div></article>
        <article className="podium-card podium-silver"><span className="podium-medal"><Medal size={27} /></span><div className="podium-copy"><span>2ND PLACE</span><strong>Runner-up</strong><h3>{runnerUpTeam?.team_name ?? '—'}</h3></div><div className="podium-points"><strong>2</strong><span>points</span></div></article>
      </div>
      <div className="result-awards"><span><b>Player of the tournament</b>{awardPlayer(tournamentResult?.player_of_tournament_id)}</span>{activePage === 'Cricket' && <><span><b>Best batsman</b>{awardPlayer(tournamentResult?.best_batsman_id)}</span><span><b>Best bowler</b>{awardPlayer(tournamentResult?.best_bowler_id)}</span></>}</div>
    </div>
  }

  async function refreshRoster() {
    const [teamResult, playerResult] = await Promise.all([
      supabase.from('teams').select('id, team_name, department, logo_url, created_at').order('team_name'),
      supabase.from('players').select('id, team_id, player_name, gender').order('player_name'),
    ])
    setTeams(teamResult.data ?? [])
    setPlayers(playerResult.data ?? [])
  }

  async function deleteRecord(table, record) {
    if (!isAdmin || !window.confirm(`Delete ${record.team_name ?? record.player_name}? This cannot be undone.`)) return
    const { error } = await supabase.from(table).delete().eq('id', record.id)
    if (error) {
      setNotice(error.message)
      return
    }
    setNotice('Record deleted.')
    refreshRoster()
  }

  async function signIn(event) {
    event.preventDefault()
    if (!isSupabaseConfigured) return
    const values = Object.fromEntries(new FormData(event.currentTarget).entries())
    const { data, error } = await supabase.auth.signInWithPassword({
      email: values.email,
      password: values.password,
    })
    if (error) {
      setNotice(error.message)
      return
    }
    if (data.user?.app_metadata?.role !== 'admin') {
      await supabase.auth.signOut()
      setNotice('This account does not have administrator access.')
      return
    }
    setDialog(null)
    setNotice('Signed in as administrator.')
  }

  async function signOut() {
    await supabase.auth.signOut()
    setNotice('Signed out. You can still view the tournament.')
  }

  function navigate(page) {
    setActivePage(page)
    setSidebarOpen(false)
    setSearchText('')
  }

  return (
    <div className="app-shell">
      <aside className={`sidebar ${sidebarOpen ? 'sidebar-open' : ''}`}>
        <div className="brand-lockup">
          <img className="brand-logo" src="/sports-carnival-logo.svg" alt="LogicValley Sports Carnival 2026" />
          <button className="icon-button sidebar-close" aria-label="Close menu" onClick={() => setSidebarOpen(false)}><X size={19} /></button>
        </div>
        <div className="nav-caption">WORKSPACE</div>
        <nav className="primary-nav" aria-label="Main navigation">
          {sports.map(({ name, icon: Icon }) => (
            <button key={name} className={`nav-item ${activePage === name ? 'nav-item-active' : ''}`} onClick={() => navigate(name)}>
              <Icon size={18} strokeWidth={1.8} /><span>{name}</span>
              {name === 'Dashboard' && <span className="nav-indicator" />}
            </button>
          ))}
        </nav>
        <div className="sidebar-bottom">
          <div className="access-note"><ShieldCheck size={17} /><span>Public viewing enabled</span></div>
          <div className="tournament-select"><span className="tournament-dot" /><span>Annual Sports Meet</span><ChevronDown size={15} /></div>
        </div>
      </aside>

      {sidebarOpen && <button className="mobile-scrim" aria-label="Close navigation" onClick={() => setSidebarOpen(false)} />}

      <main className="main-shell">
        <header className="topbar">
          <button className="icon-button menu-button" aria-label="Open menu" onClick={() => setSidebarOpen(true)}><Menu size={21} /></button>
          <div className="breadcrumbs"><span>Annual Sports Meet</span><ChevronRight size={14} /><strong>{activePage}</strong></div>
          <div className="topbar-actions">
            <span className="view-badge"><Eye size={14} /> Public view</span>
            {isAdmin ? (
              <button className="profile-button" onClick={signOut}><span className="avatar">A</span><span className="profile-name">Administrator</span><LogOut size={15} /></button>
            ) : (
              <button className="admin-login" onClick={() => setDialog({ kind: 'login' })}><LogIn size={15} /> Admin sign in</button>
            )}
          </div>
        </header>

        <section className="page-content">
          {!isSupabaseConfigured && (
            <div className="setup-banner"><CircleHelp size={17} /><span>Supabase is not connected yet. Add the public project URL and anon key to <code>.env.local</code> to load tournament data.</span></div>
          )}
          {notice && <div className="notice" role="status">{notice}<button className="icon-button" aria-label="Dismiss message" onClick={() => setNotice('')}><X size={16} /></button></div>}
          {activePage === 'Dashboard' && renderDashboard()}
          {activePage === 'Teams' && renderTeams()}
          {activePage === 'Points Table' && renderPoints()}
          {!['Dashboard', 'Teams', 'Points Table'].includes(activePage) && renderSport()}
        </section>
      </main>

      {dialog && (
        <div className="modal-backdrop" role="presentation" onMouseDown={(event) => { if (event.target === event.currentTarget) setDialog(null) }}>
          <section className="modal" role="dialog" aria-modal="true" aria-labelledby="modal-title">
            <div className="modal-heading"><div><span className="eyebrow">{dialog.kind === 'login' ? 'SECURE ACCESS' : 'ROSTER MANAGEMENT'}</span><h2 id="modal-title">{dialog.kind === 'login' ? 'Administrator sign in' : `${dialog.record ? 'Edit' : 'Add'} ${dialog.kind}`}</h2></div><button className="icon-button" aria-label="Close" onClick={() => setDialog(null)}><X size={19} /></button></div>
            {dialog.kind === 'login' ? (
              <form className="modal-form" onSubmit={signIn}>
                <p className="form-intro">Sign in with an administrator account to manage teams, players, and results. Everyone can view tournament screens without signing in.</p>
                <label>Email address<input name="email" type="email" autoComplete="username" required /></label>
                <label>Password<input name="password" type="password" autoComplete="current-password" required /></label>
                <button className="primary-button full-button" type="submit">Sign in</button>
              </form>
            ) : dialog.kind === 'pair' ? (
              <form className="modal-form" onSubmit={saveRecord}>
                <p className="form-intro">Choose two players from one team. Players must match the selected category.</p>
                <label>Team<select name="team_id" value={pairTeamId} onChange={(event) => setPairTeamId(event.target.value)} required><option value="">Select team</option>{teams.map((team) => <option key={team.id} value={team.id}>{team.team_name}</option>)}</select></label>
                <label>Player 1<select key={`player1-${dialog.record?.id ?? 'new'}-${pairTeamId}`} name="player1_id" defaultValue={dialog.record?.participant_players?.[0]?.player_id ?? ''} required><option value="">Select player</option>{availablePairPlayers.map((player) => <option key={player.id} value={player.id}>{player.player_name}</option>)}</select></label>
                <label>Player 2<select key={`player2-${dialog.record?.id ?? 'new'}-${pairTeamId}`} name="player2_id" defaultValue={dialog.record?.participant_players?.[1]?.player_id ?? ''} required><option value="">Select player</option>{availablePairPlayers.map((player) => <option key={player.id} value={player.id}>{player.player_name}</option>)}</select></label>
                <label>Pair name (optional)<input name="participant_name" defaultValue={dialog.record?.participant_name ?? ''} maxLength={120} placeholder="Defaults to player names" /></label>
                <div className="modal-actions"><button className="secondary-button" type="button" onClick={() => setDialog(null)}>Cancel</button><button className="primary-button" type="submit">{dialog.record ? 'Save pair' : 'Add pair'}</button></div>
              </form>
            ) : dialog.kind === 'participant' ? (
              <form className="modal-form" onSubmit={saveRecord}>
                <p className="form-intro">Choose one registered player from a team that matches the category.</p>
                <label>Team<select name="team_id" value={pairTeamId} onChange={(event) => setPairTeamId(event.target.value)} required><option value="">Select team</option>{teams.map((team) => <option key={team.id} value={team.id}>{team.team_name}</option>)}</select></label>
                <label>Player<select key={`single-${dialog.record?.id ?? 'new'}-${pairTeamId}`} name="player_id" defaultValue={dialog.record?.participant_players?.[0]?.player_id ?? ''} required><option value="">Select player</option>{availablePairPlayers.map((player) => <option key={player.id} value={player.id}>{player.player_name}</option>)}</select></label>
                <div className="modal-actions"><button className="secondary-button" type="button" onClick={() => setDialog(null)}>Cancel</button><button className="primary-button" type="submit">{dialog.record ? 'Save player' : 'Add player'}</button></div>
              </form>
            ) : (
              <form className="modal-form" onSubmit={saveRecord}>
                {dialog.kind === 'team' ? <>
                  <label>Team name<input name="team_name" defaultValue={dialog.record?.team_name ?? ''} required maxLength={80} /></label>
                </> : <>
                  <label>Player name<input name="player_name" defaultValue={dialog.record?.player_name ?? ''} required maxLength={80} /></label>
                  <label>Team<select name="team_id" defaultValue={dialog.record?.team_id ?? dialog.teamId ?? ''} required><option value="" disabled>Select team</option>{teams.map((team) => <option key={team.id} value={team.id}>{team.team_name}</option>)}</select></label>
                  <label>Gender<select name="gender" defaultValue={dialog.record?.gender ?? ''} required><option value="" disabled>Select gender</option><option>Male</option><option>Female</option></select></label>
                </>}
                <div className="modal-actions"><button className="secondary-button" type="button" onClick={() => setDialog(null)}>Cancel</button><button className="primary-button" type="submit">Save changes</button></div>
              </form>
            )}
          </section>
        </div>
      )}
    </div>
  )

  function renderDashboard() {
    const sportsCount = points.reduce((set, row) => set.add(row.sport_id), new Set()).size
    const stats = [
      { label: 'Registered teams', value: teams.length, icon: Users, tone: 'green' },
      { label: 'Players', value: players.length, icon: Activity, tone: 'blue' },
      { label: 'Sports with results', value: sportsCount, icon: Trophy, tone: 'gold' },
      { label: 'Points awarded', value: points.reduce((sum, row) => sum + (row.total_points ?? 0), 0), icon: Award, tone: 'coral' },
    ]
    return <>
      <div className="page-heading"><div><span className="eyebrow">TOURNAMENT OVERVIEW</span><h1>Good day, sports fans<span className="heading-period">.</span></h1><p>Follow the teams, fixtures, and results from one place.</p></div><div className="dashboard-header-actions"><button className="secondary-button" onClick={downloadAllSchedulesExcel} disabled={loading || isExportingSchedules}><Download size={15} /> {isExportingSchedules ? 'Exporting…' : 'Download schedules'}</button><span className="season-pill"><span /> 2026 Season</span></div></div>
      <div className="stats-grid">{stats.map(({ label, value, icon: Icon, tone }) => <article className="stat-card" key={label}><div className={`stat-icon ${tone}`}><Icon size={19} /></div><div className="stat-label">{label}</div><div className="stat-value">{loading ? '—' : value}</div></article>)}</div>
      <div className="dashboard-grid">
        <section className="surface standings-surface"><div className="section-heading"><div><span className="eyebrow">LEADERBOARD</span><h2>Overall points</h2></div><button className="text-button" onClick={() => navigate('Points Table')}>View table <ChevronRight size={15} /></button></div>{renderStandings(totalPoints.slice(0, 5))}</section>
        <section className="feature-panel"><div className="feature-top"><span className="live-dot" /> TOURNAMENT HUB</div><div className="feature-art"><div className="art-orbit orbit-one" /><div className="art-orbit orbit-two" /><Trophy size={70} strokeWidth={1.2} /></div><h2>Every match<br />has a story.</h2><p>Browse each sport's tournament page for brackets and published results.</p><button className="feature-link" onClick={() => navigate('Badminton')}>Explore tournaments <ChevronRight size={16} /></button></section>
      </div>
      <section className="surface activity-surface"><div className="section-heading"><div><span className="eyebrow">ROSTER SNAPSHOT</span><h2>Teams in the tournament</h2></div><button className="text-button" onClick={() => navigate('Teams')}>All teams <ChevronRight size={15} /></button></div>{teams.length ? <div className="team-strip">{teams.slice(0, 6).map((team, index) => <div className="team-chip" key={team.id}><span className={`team-monogram monogram-${index % 5}`}>{team.team_name.slice(0, 1).toUpperCase()}</span><div><strong>{team.team_name}</strong><small>{players.filter((player) => player.team_id === team.id).length} players</small></div></div>)}</div> : <div className="empty-inline">{loading ? 'Loading roster…' : 'Team entries will appear here once they are registered.'}</div>}</section>
    </>
  }

  function renderTeams() {
    const preferredTeamNames = [
      ['alapparai squad'],
      ['mgic mushrooms', 'magic mushrooms', 'magic mashrooms'],
      ['moorthy anna gang'],
      ['yathuku', 'edhuku'],
    ]
    const orderedFilteredTeams = [...filteredTeams].sort((first, second) => {
      const firstName = first.team_name.trim().toLowerCase()
      const secondName = second.team_name.trim().toLowerCase()
      const firstOrder = preferredTeamNames.findIndex((names) => names.includes(firstName))
      const secondOrder = preferredTeamNames.findIndex((names) => names.includes(secondName))
      return (firstOrder < 0 ? Infinity : firstOrder) - (secondOrder < 0 ? Infinity : secondOrder)
    })
    return <>
      <div className="page-heading compact-heading"><div><span className="eyebrow">TOURNAMENT ROSTER</span><h1>Teams & players<span className="heading-period">.</span></h1><p>Browse registered teams and their player rosters.</p></div>{isAdmin && <button className="primary-button" onClick={() => setDialog({ kind: 'team' })}><Plus size={17} /> Add team</button>}</div>
      <section className="surface teams-toolbar"><div className="roster-count"><strong>{teams.length}</strong> teams <span className="count-divider">/</span> {players.length} players</div><label className="search-box"><Search size={16} /><input placeholder="Search teams or players" value={searchText} onChange={(event) => setSearchText(event.target.value)} /></label></section>
      {loading ? <div className="empty-state">Loading the tournament roster…</div> : filteredTeams.length === 0 ? <section className="surface empty-state"><span className="empty-icon"><Users size={24} /></span><strong>{teams.length ? 'No matching teams or players' : 'No teams registered yet'}</strong><p>{teams.length ? 'Try another search.' : isAdmin ? 'Add the first team to start building the tournament roster.' : 'The tournament roster will appear here once teams are registered.'}</p>{isAdmin && !teams.length && <button className="primary-button" onClick={() => setDialog({ kind: 'team' })}><Plus size={16} /> Add first team</button>}</section> : <div className="team-roster-grid">{orderedFilteredTeams.map((team) => {
        const teamPlayers = players.filter((player) => player.team_id === team.id && `${player.player_name} ${player.gender}`.toLowerCase().includes(searchText.toLowerCase()))
        if (searchText && !teamPlayers.length && !team.team_name.toLowerCase().includes(searchText.toLowerCase())) return null
        return <article className="team-roster-card" key={team.id}>
          <header className="team-roster-heading"><h2>{team.team_name}</h2><span>{teamPlayers.length} {teamPlayers.length === 1 ? 'player' : 'players'}</span>{isAdmin && <div className="row-actions"><button onClick={() => setDialog({ kind: 'team', record: team })}>Edit</button><button className="danger-action" onClick={() => deleteRecord('teams', team)}>Delete</button></div>}</header>
          <div className="team-roster-table-wrap"><table className="team-roster-table"><thead><tr><th>S.No</th><th>Name</th><th>Gender</th>{isAdmin && <th>Actions</th>}</tr></thead><tbody>{teamPlayers.length ? teamPlayers.map((player, index) => <tr key={player.id}><td>{index + 1}</td><td>{player.player_name}</td><td>{player.gender}</td>{isAdmin && <td><div className="row-actions"><button onClick={() => setDialog({ kind: 'player', record: player })}>Edit</button><button className="danger-action" onClick={() => deleteRecord('players', player)}>Delete</button></div></td>}</tr>) : <tr><td className="team-no-players" colSpan={isAdmin ? 4 : 3}>No players registered</td></tr>}</tbody></table></div>
          {isAdmin && <button className="add-player-button" onClick={() => setDialog({ kind: 'player', teamId: team.id })}><Plus size={13} /> Add player</button>}
        </article>
      })}</div>}
      <div className="permission-note"><Eye size={16} /><span>Anyone can view this roster. Only signed-in administrators can add, edit, or delete teams and players.</span>{!isAdmin && <button onClick={() => setDialog({ kind: 'login' })}>Admin sign in <ChevronRight size={14} /></button>}</div>
    </>
  }

  function renderPoints() {
    const filteredRows = filterSport === 'Overall' ? points : points.filter((row) => row.sports?.sport_name === filterSport)
    const categoryGroups = new Map()
    allCategories
      .filter((category) => filterSport === 'Overall' || category.sports?.sport_name === filterSport)
      .forEach((category) => categoryGroups.set(`${category.sports?.sport_name ?? 'Sport'}:${category.category_name}`, {
        sportName: category.sports?.sport_name ?? 'Sport',
        categoryName: category.category_name,
        rows: new Map(teams.map((team) => [team.id, {
          team_name: team.team_name,
          winner_points: 0,
          runner_up_points: 0,
          total_points: 0,
        }])),
      }))
    filteredRows.forEach((row) => {
      const sportName = row.sports?.sport_name ?? 'Sport'
      const categoryName = row.tournament_categories?.category_name ?? 'Category'
      const groupKey = `${sportName}:${categoryName}`
      if (!categoryGroups.has(groupKey)) categoryGroups.set(groupKey, {
        sportName,
        categoryName,
        rows: new Map(teams.map((team) => [team.id, {
          team_name: team.team_name,
          winner_points: 0,
          runner_up_points: 0,
          total_points: 0,
        }])),
      })
      const group = categoryGroups.get(groupKey)
      const teamName = row.teams?.team_name ?? 'Team'
      const teamKey = row.team_id ?? teamName
      const current = group.rows.get(teamKey) ?? { team_name: teamName, winner_points: 0, runner_up_points: 0, total_points: 0 }
      current.winner_points += row.winner_points ?? 0
      current.runner_up_points += row.runner_up_points ?? 0
      current.total_points += row.total_points ?? 0
      group.rows.set(teamKey, current)
    })
    const combinedRows = filterSport === 'Overall' ? totalPoints : [...filteredRows.reduce((totals, row) => {
      const teamName = row.teams?.team_name ?? 'Team'
      const teamKey = row.team_id ?? teamName
      const current = totals.get(teamKey) ?? { team_name: teamName, winner_points: 0, runner_up_points: 0, total_points: 0 }
      current.winner_points += row.winner_points ?? 0
      current.runner_up_points += row.runner_up_points ?? 0
      current.total_points += row.total_points ?? 0
      totals.set(teamKey, current)
      return totals
    }, new Map(teams.map((team) => [team.id, {
      team_name: team.team_name,
      winner_points: 0,
      runner_up_points: 0,
      total_points: 0,
    }]))).values()].sort((a, b) => b.total_points - a.total_points || a.team_name.localeCompare(b.team_name))
    return <>
      <div className="page-heading compact-heading"><div><span className="eyebrow">RESULTS & STANDINGS</span><h1>Points table<span className="heading-period">.</span></h1><p>Each sport and category has its own standings table.</p></div><label className="select-filter"><span>SPORT</span><select value={filterSport} onChange={(event) => setFilterSport(event.target.value)}><option>Overall</option>{sports.slice(2, 9).map((sport) => <option key={sport.name}>{sport.name}</option>)}</select></label></div>
      <section className="surface standings-surface full-standings"><div className="section-heading"><div><span className="eyebrow">{filterSport.toUpperCase()}</span><h2>{filterSport === 'Overall' ? 'All sports combined' : `${filterSport} combined`}</h2></div><span className="points-legend"><b>4</b> Winner <b>2</b> Runner-up</span></div>{renderStandings(combinedRows)}</section>
      <div className="category-points-grid">{[...categoryGroups.values()].map((group) => <section className="surface standings-surface full-standings" key={`${group.sportName}:${group.categoryName}`}><div className="section-heading"><div><span className="eyebrow">{group.sportName.toUpperCase()}</span><h2>{group.categoryName}</h2></div></div>{renderStandings([...group.rows.values()].sort((a, b) => b.total_points - a.total_points))}</section>)}</div>
      {!loading && categoryGroups.size === 0 && <div className="surface empty-state"><span className="empty-icon"><Award size={24} /></span><strong>No confirmed results yet</strong><p>Points will appear in each sport’s table after a winner and runner-up are saved.</p></div>}
    </>
  }

  function renderStandings(rows) {
    if (loading) return <div className="empty-inline">Loading points table…</div>
    if (!rows.length) return <div className="empty-inline">No teams registered yet.</div>
    return <div className="points-table-scroll"><table className="points-data-table"><thead><tr><th>S.No</th><th>Team name</th><th>Winner</th><th>Runner-up</th><th>Total</th></tr></thead><tbody>{rows.map((row, index) => <tr key={row.team_name}><td>{index + 1}</td><td><strong>{row.team_name}</strong></td><td>{row.winner_points ?? 0}</td><td>{row.runner_up_points ?? 0}</td><td><strong>{row.total_points}</strong></td></tr>)}</tbody></table></div>
  }

  function renderSport() {
    const isBracketSport = ['Badminton', 'Table Tennis', 'Carrom', 'Chess'].includes(activePage)
    const supportsPairs = ['Badminton', 'Carrom', 'Table Tennis'].includes(activePage)
    const isTeamResultSport = ['Volleyball', 'Football', 'Cricket'].includes(activePage)
    const supportsKnockoutRoster = ['Badminton', 'Carrom', 'Table Tennis', 'Chess'].includes(activePage)
    const showPairManager = supportsPairs && activePairCategory?.competition_type === 'doubles'
    const showSingleManager = supportsKnockoutRoster && activePairCategory?.competition_type === 'singles'
    const canManageKnockoutRoster = showPairManager || showSingleManager
    return <>
      <div className="page-heading compact-heading"><div><span className="eyebrow">TOURNAMENT VIEWER</span><h1>{activePage}<span className="heading-period">.</span></h1><p>Published fixtures and results are visible to everyone.</p></div><span className="season-pill"><Eye size={14} /> Public view</span></div>
      {isTeamResultSport && <section className="surface result-management-surface"><div className="section-heading"><div><span className="eyebrow">FINAL RESULTS</span><h2>{activePage} result</h2></div><label className="select-filter"><span>CATEGORY</span><select value={categoryId} onChange={(event) => setCategoryId(event.target.value)}><option value="">Select category</option>{pairCategories.map((category) => <option key={category.id} value={category.id}>{category.category_name}</option>)}</select></label></div>
        {isAdmin && (!tournamentResult || isEditingResult) ? <form className="result-form" key={`${categoryId}-${tournamentResult?.winner_team_id ?? 'new'}-${tournamentResult?.runner_up_team_id ?? 'new'}`} onSubmit={saveTeamSportResult}>
          <label>Tournament winner<select name="winner_team_id" defaultValue={tournamentResult?.winner_team_id ?? ''} required><option value="">Select winner</option>{teams.map((team) => <option key={team.id} value={team.id}>{team.team_name}</option>)}</select></label>
          <label>Runner-up<select name="runner_up_team_id" defaultValue={tournamentResult?.runner_up_team_id ?? ''} required><option value="">Select runner-up</option>{teams.map((team) => <option key={team.id} value={team.id}>{team.team_name}</option>)}</select></label>
          <label>Player of the tournament<select name="player_of_tournament_id" defaultValue={tournamentResult?.player_of_tournament_id ?? ''} required><option value="">Select player</option>{players.map((player) => <option key={player.id} value={player.id}>{player.player_name} ({teams.find((team) => team.id === player.team_id)?.team_name ?? 'Team'})</option>)}</select></label>
          {activePage === 'Cricket' && <>
            <label>Best batsman<select name="best_batsman_id" defaultValue={tournamentResult?.best_batsman_id ?? ''} required><option value="">Select player</option>{players.map((player) => <option key={player.id} value={player.id}>{player.player_name} ({teams.find((team) => team.id === player.team_id)?.team_name ?? 'Team'})</option>)}</select></label>
            <label>Best bowler<select name="best_bowler_id" defaultValue={tournamentResult?.best_bowler_id ?? ''} required><option value="">Select player</option>{players.map((player) => <option key={player.id} value={player.id}>{player.player_name} ({teams.find((team) => team.id === player.team_id)?.team_name ?? 'Team'})</option>)}</select></label>
          </>}
          <button className="primary-button" type="submit" disabled={!categoryId}><Trophy size={15} /> Save result</button>
        </form> : tournamentResult ? <div className="saved-result-summary"><strong>Result saved</strong><span>{teams.find((team) => team.id === tournamentResult.winner_team_id)?.team_name ?? '—'} is winner</span><span>{teams.find((team) => team.id === tournamentResult.runner_up_team_id)?.team_name ?? '—'} is runner-up</span>{isAdmin && <button className="text-button" onClick={() => setIsEditingResult(true)}>Update result <ChevronRight size={14} /></button>}</div> : <div className="empty-inline">No final result has been saved for this category.</div>}
      </section>}
      {supportsKnockoutRoster && <section className="surface pair-management-surface"><div className="section-heading"><div><span className="eyebrow">TOURNAMENT ENTRIES</span><h2>{showPairManager ? 'Registered pairs' : 'Registered players'}</h2></div><div className="pair-toolbar">{!isAdmin && ['Badminton', 'Carrom', 'Chess'].includes(activePage) && <button className="admin-login" onClick={() => setDialog({ kind: 'login' })}><LogIn size={15} /> Admin sign in</button>}<label className="select-filter"><span>CATEGORY</span><select value={categoryId} onChange={(event) => setCategoryId(event.target.value)}><option value="">Select category</option>{pairCategories.filter((category) => ['doubles', 'singles'].includes(category.competition_type)).map((category) => <option key={category.id} value={category.id}>{category.category_name}</option>)}</select></label>{isAdmin && canManageKnockoutRoster && (showPairManager || !fixtureMatches.length) && <button className="primary-button" onClick={() => (showPairManager ? openPairDialog() : openParticipantDialog())}><Plus size={16} /> {showPairManager ? 'Add pair' : 'Add player'}</button>}{isAdmin && canManageKnockoutRoster && fixtureMatches.length === 0 && <button className="secondary-button" onClick={generateFixtures} disabled={pairRows.length < 2}><Trophy size={15} /> Generate fixtures</button>}{isAdmin && canManageKnockoutRoster && fixtureMatches.length > 0 && <button className="secondary-button danger-button" onClick={deleteFixtures}><Trash2 size={15} /> Delete bracket</button>}</div></div>
        {!activePairCategory ? <div className="empty-inline">{isSupabaseConfigured ? 'No tournament categories found. Run the Supabase seed script first.' : 'Connect Supabase to load tournament data.'}</div> : !canManageKnockoutRoster ? <div className="empty-inline">Select a singles or doubles category.</div> : pairRows.length === 0 ? <div className="empty-state"><span className="empty-icon"><Users size={24} /></span><strong>No {showPairManager ? 'pairs' : 'players'} registered</strong><p>{isAdmin ? showPairManager ? 'Add players to a team first, then register a pair here.' : 'Register a player from the Teams roster.' : 'Registered tournament entries will appear here.'}</p>{isAdmin && (showPairManager || !fixtureMatches.length) && <button className="primary-button" onClick={() => (showPairManager ? openPairDialog() : openParticipantDialog())}><Plus size={16} /> Add first {showPairManager ? 'pair' : 'player'}</button>}</div> : <div className="table-scroll"><table className="data-table"><thead><tr><th>{showPairManager ? 'PAIR' : 'PLAYER'}</th><th>TEAM</th>{showPairManager && <><th>PLAYER 1</th><th>PLAYER 2</th></>}{isAdmin && <th className="action-column">ACTIONS</th>}</tr></thead><tbody>{pairRows.map((participant) => <tr key={participant.id}><td><strong className="pair-name-cell">{participant.participant_name}</strong></td><td>{participant.teams?.team_name ?? '—'}</td>{showPairManager && <><td>{participant.participant_players?.[0]?.players?.player_name ?? '—'}</td><td>{participant.participant_players?.[1]?.players?.player_name ?? '—'}</td></>}{isAdmin && <td><div className="row-actions"><button onClick={() => (showPairManager ? openPairDialog(participant) : openParticipantDialog(participant))}>Edit</button><button className="danger-action" onClick={() => deletePair(participant)}>Delete</button></div></td>}</tr>)}</tbody></table></div>}
        <div className="pair-access-note"><Eye size={15} /><span>{showPairManager ? 'Pairs' : 'Players'} are visible to everyone. Only the administrator can change them.</span></div>
      </section>}
      <section className="surface sport-view-surface"><div className="section-heading"><div><span className="eyebrow">{isBracketSport ? 'KNOCKOUT FIXTURES' : 'TOURNAMENT RESULTS'}</span><h2>{isBracketSport ? 'Tournament bracket' : 'Final standings'}</h2></div>{isBracketSport && fixtureMatches.length > 0 && <button className="secondary-button" onClick={downloadFixturesExcel}><Download size={15} /> Download Excel</button>}{isAdmin && !supportsKnockoutRoster && <span className="admin-only-tag"><ShieldCheck size={14} /> Admin controls</span>}</div>
        {isBracketSport && fixtureMatches.length > 0 ? renderBracket() : isTeamResultSport && tournamentResult ? renderTeamFinalStandings() : <div className="empty-state"><span className="empty-icon"><Award size={24} /></span><strong>{isBracketSport ? 'No fixtures generated yet' : 'No results published yet'}</strong><p>{isBracketSport && isAdmin && canManageKnockoutRoster ? `Register at least two ${showPairManager ? 'pairs' : 'players'}, then generate fixtures above.` : 'Published tournament data will appear here.'}</p></div>}
        <div className="fixture-footnote"><span className="live-dot" /> Completed match winners advance automatically to the next round.</div>
      </section>
      <div className="permission-note"><Eye size={16} /><span>Fixtures and results are public. Administrative changes are restricted to signed-in administrators.</span></div>
    </>
  }

  function renderBracket() {
    const teamColorById = new Map(teams.map((team, index) => [team.id, bracketTeamColors[index % bracketTeamColors.length]]))
    const roundNumbers = [...new Set(fixtureMatches.map((match) => match.round_number))].sort((a, b) => a - b)
    const firstRoundCount = fixtureMatches.filter((match) => match.round_number === roundNumbers[0]).length
    const boardHeight = Math.max(224, firstRoundCount * 112)
    const finalRoundNumber = roundNumbers.at(-1)
    const finalMatch = fixtureMatches.find((match) => match.round_number === finalRoundNumber)
    const champion = [finalMatch?.participant1, finalMatch?.participant2].find((participant) => participant?.id === finalMatch?.winner_participant_id)
    return <div className="bracket-scroll"><div className="bracket-board" style={{ minHeight: `${boardHeight}px` }}>
      {roundNumbers.map((roundNumber, roundIndex) => {
        const roundMatches = fixtureMatches.filter((match) => match.round_number === roundNumber)
        const step = 112 * (2 ** roundIndex)
        const roundLabel = roundNumber === finalRoundNumber ? 'FINAL'
          : roundNumber === finalRoundNumber - 1 ? 'SEMI-FINALS'
            : roundNumber === finalRoundNumber - 2 ? 'QUARTER-FINALS'
              : `ROUND ${roundNumber}`
        return <section className="bracket-round-column" key={roundNumber} style={{ height: `${boardHeight}px` }}>
          <h3>{roundLabel}</h3>
          {roundMatches.map((match, matchIndex) => {
            const top = step / 2 - 38 + matchIndex * step
            const participants = [match.participant1, match.participant2]
            const participantIds = [match.participant1_id, match.participant2_id]
            const isByeMatch = match.match_status === 'bye'
              || (match.match_status === 'completed' && !match.participant2_id && match.winner_participant_id === match.participant1_id)
            const connectionClass = roundIndex < roundNumbers.length - 1
              ? matchIndex % 2 === 0 ? 'bracket-pair-top' : 'bracket-pair-bottom'
              : ''
            return <div className={`bracket-match-position ${connectionClass}`} key={match.id} style={{ top: `${top}px`, '--stem-height': `${step / 2}px` }}>
              <article className="bracket-match-card">
                {participants.map((participant, slotIndex) => {
                  const participantId = participantIds[slotIndex]
                  const isWinner = participantId && participantId === match.winner_participant_id
                  const teamColor = participant?.team_id ? teamColorById.get(participant.team_id) : null
                  return <div className={`bracket-participant ${teamColor ? 'bracket-participant-team' : ''} ${isWinner ? 'bracket-winner' : ''}`} key={`${match.id}-${slotIndex}`} style={teamColor ? { '--team-accent': teamColor.accent, '--team-tint': teamColor.tint } : undefined}>
                    <span className="bracket-team-name">{participant?.teams?.team_name ?? (isByeMatch ? 'BYE' : 'TBD')}</span>
                    <strong>{participant?.participant_name ?? (isByeMatch ? 'BYE' : 'Waiting for opponent')}</strong>
                  </div>
                })}
                <div className="bracket-match-footer">
                  <span>{isByeMatch ? match.winner_participant_id ? 'BYE · ADVANCED' : 'BYE · SELECT TO ADVANCE' : match.match_status === 'completed' ? 'COMPLETED' : `MATCH ${match.match_number}`}</span>
                  {(() => {
                    return isAdmin && participants[0] && (isByeMatch || participants[1]) && <select aria-label={`Winner for match ${match.match_number}`} value={match.winner_participant_id ?? ''} onChange={(event) => event.target.value === '__incomplete__' ? setMatchIncomplete(match) : setMatchWinner(match, event.target.value)}><option value="" disabled>{isByeMatch ? 'Select advancement' : 'Set winner'}</option><option value={participants[0].id}>{isByeMatch ? `Advance ${participants[0].participant_name} (BYE)` : participants[0].participant_name}</option>{participants[1] && <option value={participants[1].id}>{participants[1].participant_name}</option>}{match.match_status === 'completed' && <option value="__incomplete__">Mark incomplete</option>}</select>
                  })()}
                </div>
              </article>
              {roundIndex < roundNumbers.length - 1 && matchIndex % 2 === 0 && <span className="bracket-entry-line" style={{ top: `${52 + step / 2}px` }} />}
            </div>
          })}
        </section>
      })}
      <div className="bracket-champion"><Award size={28} /><span>CHAMPION</span><strong>{champion?.participant_name ?? 'To be determined'}</strong><small>{champion?.teams?.team_name ?? ''}</small></div>
    </div></div>
  }
}

export default App