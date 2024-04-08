export function lst(workbook) {
    const sectionUnitNRI = {}

    const WDTable = [
        'PRSPNAME',
        'LSEAGMT',
        'LSENAME',
        'EXPDATE',
        'TSTATUS',
        'LSETYPE',
        'TLEGAL',
        'DEPTIVL',
        'LSEROY',
        'NMA',
        'NRI',
        'TSTATUS']
    const result = { rawTable: [], sectionGross: {}, formations: [] }
    const worksheet1 = workbook.sheet(workbook.sheetNames[0], { headers: true })
    workbook.sheet(workbook.sheetNames[1], { headers: true }).forEach(formation => {
        if (formation['Formation Name'] === undefined) return
        let selectedFormation = { name: formation['Formation Name'], depthIntervals: [] }
        Object.keys(formation).forEach(column => {
            if (column !== 'Formation Name' && formation[column] === column) selectedFormation.depthIntervals.push(column)
        })
        result.formations.push(selectedFormation)
    })

    let resultRow = {}
    const includedSectionIds = {}
    worksheet1.forEach((row, i) => {
        resultRow = { STR: `${row['SECTION']}-${row["TWNSHIP"]}${row["RANGE"]}I` }

        if (row["TRACTID"] == '' || row["TRACTID"] == undefined || row["TSTATUS"] == 'SOLD' || row["TSTATUS"] == 'EXPIRED' || row["SECTION"] == "" || row["SECTION"] == undefined || row["TWNSHIP"] == undefined) return
        let prospect = row["PRSPNAME"]
        includedSectionIds[sectionToId(resultRow.STR)] = 1
        if (typeof (row['EXPDATE']) === 'number') row['EXPDATE'] = new Date((row['EXPDATE'] - (25567 + 2)) * 86400000)
        row['TLEGAL'] = row['TLEGAL'].trim()
        row['LSENAME'] = row['LSENAME'].trim()
        if (row['TLEGAL'].length > 20) row['TLEGAL'] = `${row['TLEGAL'].substring(0, 27)}...`
        if (row['LSENAME'].length > 20) row['LSENAME'] = `${row['LSENAME'].substring(0, 27)}...`
        WDTable.forEach(header => {
            resultRow[header] = row[header]
        })
        resultRow.NMA = parseFloat(row['NETACRES'])
        result.rawTable.push(resultRow)
        /*
                let id = `${row['LSEAGMT']}${row['TRACTID']}${row['TLEGAL']}${row['DEPTIVL']}`
                if (ids[id] !== undefined) { console.log(`Duplicate Row ${id}`); return }
                ids[id] = 1
        
                let leaseID = row['LSEAGMT']
                data[leaseID] = data[leaseID] || { tracts: [], NMA: 0, depths: [], STR: `${row['SECTION']}-${row["TWNSHIP"]}${row["RANGE"]}` }
                
                let legal = `${row['TLEGAL']}(${row['NETACRES']})`
                if (data[leaseID].tracts.indexOf(legal) < 0) {
                    data[leaseID].tracts.push(legal)
                    data[leaseID].NMA += parseFloat(row['NETACRES'])
                }
                let depth = row['DEPTIVL']
        
                if (data[leaseID].depths.indexOf(depth) < 0) data[leaseID].depths.push(depth)
                data[leaseID].row = row
        
            })
        
            Object.keys(data).forEach(lease => {
                let current = data[lease]
                let cLease = {}
                WDTable.forEach(header => {
                    cLease[header] = current.row[header]
                })
                cLease.STR = current.STR,
                    cLease['DEPTIVL'] = current.depths.join(' | '),
                    cLease['TLEGAL'] = current.tracts.join(' | '),
                    cLease['NMA'] = current.NMA
                result.push(cLease)
                */
    })

    if (workbook.sheetNames[2] !== undefined) workbook.sheet(workbook.sheetNames[2], { headers: true }).forEach((v, i) => {

        let id = sectionToId(v["Section"])
        if (id === null) return
        if (includedSectionIds[id] === 1) result.sectionGross[id] = parseFloat(v["Gross"])

    })
    Object.keys(includedSectionIds).forEach(id => {
        if (result.sectionGross[id] === undefined) result.sectionGross[id] = 640
    })
    return result
}
export function sectionToId(section) {
    if (section === undefined || section.indexOf === undefined) return null
    if (section.indexOf('-') < 0) {
        if (section.length > 15) console.log(`Section is too long ${section}`)
        const sectionRegex = /(\d+)([N|S])(\d+)([E|W])(\d+)(.*)/
        let d = section.replace(sectionRegex, '$1-$2-$3-$4-$5-$6').split('-')
        while (d[0].length < 3) d[0] = `0${d[0]}`
        while (d[2].length < 3) d[2] = `0${d[2]}`
        while (d[4].length < 2) d[4] = `0${d[4]}`
        d[1] = d[1] === 'N' ? 0 : 1
        d[3] = d[3] === 'E' ? 0 : 1
        d[5] = (d[5].indexOf('I') == 0 || d[5] === '') ? 0 : 1
        return parseInt(`1${d[0]}${d[1]}1${d[2]}${d[3]}${d[4]}${d[5]}`)
    }
    if (section.length > 15) console.log(`Section is too long ${section}`)
    const sectionRegex = /(\d+)-(\d+)([N|S])(\d+)([E|W])(.*)/
    let d = section.replace(sectionRegex, '$2-$3-$4-$5-$1-$6').split('-')
    while (d[0].length < 3) d[0] = `0${d[0]}`
    while (d[2].length < 3) d[2] = `0${d[2]}`
    while (d[4].length < 2) d[4] = `0${d[4]}`
    d[1] = d[1] === 'N' ? 0 : 1
    d[3] = d[3] === 'E' ? 0 : 1
    d[5] = (d[5].indexOf('I') == 0 || d[5] === '') ? 0 : 1
    return parseInt(`1${d[0]}${d[1]}1${d[2]}${d[3]}${d[4]}${d[5]}`)
}
export function idToSection(id, sFirst) {
    id = typeof (id) !== 'string' ? id.toString() : id
    const idRegex = /1(\d\d\d)(\d)1(\d\d\d)(\d)(\d\d)(\d)/
    let d = id.replace(idRegex, '$1-$2-$3-$4-$5-$6').split('-')
    d[0] = parseInt(d[0])
    d[1] = d[1] === '0' ? 'N' : 'S'
    d[2] = parseInt(d[2])
    d[3] = d[3] === '0' ? 'E' : 'W'
    d[4] = parseInt(d[4])
    d[5] = d[5] === '0' ? "IM" : "CM"
    if (sFirst) return [d[4], '-', d[0], d[1], d[2], d[3]].join('')
    return d.join('')
}


