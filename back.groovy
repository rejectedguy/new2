<%
EnableLog("myLog", false);
result = new Object();
result.SetProperty('debug', []);
result.SetProperty('structure', []);
result.SetProperty('filtered_collaborators', []);
result.GetOptProperty('debug').push('INIT: result created at start');
try {
    subNameMap = [];
    hiddenPrefixes = ['Обособленное подразделение', 'Основное подразделение'];
notOrgIds = [];
// Загружаем переменную sort_orgs для сортировки организаций
sortOrgMap = new Object();
try {
    result.GetOptProperty('debug').push('START: loading sort_orgs');
    arrVars = tools.xquery("for $elem in custom_web_templates where $elem/code = 'testSd' return $elem");

    if (ArrayCount(arrVars) > 0) {
        for (elem in arrVars) {
            try {
                elemOpen = tools.open_doc(elem.id);
                if (elemOpen == undefined) continue;

                elemTop = elemOpen.TopElem.wvars;
                if (elemTop == undefined) continue;

                for (wvar in elemTop) {
                    if (wvar.name == 'sort_orgs') {
                        result.GetOptProperty('debug').push('sort_orgs found! value="' + String(wvar.value) + '"');
                        sortOrgValue = String(wvar.value);

                        if (sortOrgValue == '') {
                            result.GetOptProperty('debug').push('sort_orgs empty - skip');
                            continue;
                        }

                        sortOrgArray = ParseJson(sortOrgValue);
                        result.GetOptProperty('debug').push('ParseJson done: sortOrgArray count=' + ArrayCount(sortOrgArray));

                        for (item in sortOrgArray) {
                            try {
                                orgId = OptInt(String(item.__value));
                                sortOrder = String(item.comment);

                                if (orgId != 0 && sortOrder != '') {
                                    sortOrgMap.SetProperty(String(orgId), OptInt(sortOrder));
                                    result.GetOptProperty('debug').push('Sort org ID: ' + orgId + ' = order ' + sortOrder);
                                }
                            } catch(itemErr) {
                                result.GetOptProperty('debug').push('ERROR parsing sort_orgs item: ' + String(itemErr));
                            }
                        }
                    }
                }
            } catch(docErr) {
                result.GetOptProperty('debug').push('ERROR loading sort_orgs doc: ' + String(docErr));
            }
        }
    }

    result.GetOptProperty('debug').push('END: total sortOrgMap entries=' + sortOrgMap);
} catch(sortOrgErr) {
    result.GetOptProperty('debug').push('ERROR loading sort_orgs: ' + String(sortOrgErr));
}
// Загружаем ВСЕ переменные sort_org* для сортировки подразделений внутри организаций
sortOrgSubdivisionsMap = new Object(); // org_name -> { subdivision_name -> sort_order }

try {
    result.GetOptProperty('debug').push('START: loading sort_org* variables');
    arrVars = tools.xquery("for $elem in custom_web_templates where $elem/code = 'testSd' return $elem");

    if (ArrayCount(arrVars) > 0) {
        for (elem in arrVars) {
            try {
                elemOpen = tools.open_doc(elem.id);
                if (elemOpen == undefined) continue;

                elemTop = elemOpen.TopElem.wvars;
                if (elemTop == undefined) continue;

                for (wvar in elemTop) {
                    var_name = String(wvar.name);

                    // Проверяем что это переменная вида sort_org2, sort_org3, sort_org4 и т.д.
                    if (StrBegins(var_name, 'sort_org', false) && var_name != 'sort_org') {
                        // Извлекаем номер (2, 3, 4...)
                        org_number = StrRightRangePos(var_name, 8); // После "sort_org"

                        if (org_number == '' || OptInt(org_number) == 0) continue;

                        org_description = String(wvar.description); // Название организации из description

                        if (org_description == '') {
                            result.GetOptProperty('debug').push('SKIP: ' + var_name + ' - no description');
                            continue;
                        }

                        result.GetOptProperty('debug').push('Found ' + var_name + ' for org: "' + org_description + '"');

                        // Создаем карту для этой организации если еще нет
                        if (!sortOrgSubdivisionsMap.HasProperty(org_description)) {
                            sortOrgSubdivisionsMap.SetProperty(org_description, new Object());
                        }

                        org_map = sortOrgSubdivisionsMap.GetOptProperty(org_description);

                        // Загружаем entries (подразделения и их порядок)
                        if (wvar.entries == undefined) continue;

                        for (entry in wvar.entries) {
                            try {
                                sub_name = String(entry.id); // Название подразделения
                                sort_order = String(entry.name); // Порядок сортировки

                                if (sub_name != '' && sort_order != '') {
                                    org_map.SetProperty(StrLowerCase(Trim(sub_name)), OptInt(sort_order));
                                    result.GetOptProperty('debug').push('  Subdivision "' + sub_name + '" = order ' + sort_order);
                                }
                            } catch(entryErr) {
                                result.GetOptProperty('debug').push('ERROR parsing entry: ' + String(entryErr));
                            }
                        }
                    }
                }
            } catch(docErr) {
                result.GetOptProperty('debug').push('ERROR loading doc: ' + String(docErr));
            }
        }
    }

    result.GetOptProperty('debug').push('END: loaded sort_org* for organizations');
} catch(sortOrgErr) {
    result.GetOptProperty('debug').push('ERROR loading sort_org*: ' + String(sortOrgErr));
}
try {
    result.GetOptProperty('debug').push('START: loading not_org');
    arrVars = tools.xquery("for $elem in custom_web_templates where $elem/code = 'testSd' return $elem");

    if (ArrayCount(arrVars) > 0) {
        for (elem in arrVars) {
            try {
                elemOpen = tools.open_doc(elem.id);
                if (elemOpen == undefined) continue;

                elemTop = elemOpen.TopElem.wvars;
                if (elemTop == undefined) continue;

                for (wvar in elemTop) {
                    if (wvar.name == 'not_org') {
                        result.GetOptProperty('debug').push('not_org found! value="' + String(wvar.value) + '"');
                        notOrgValue = String(wvar.value);

                        if (notOrgValue == '') {
                            result.GetOptProperty('debug').push('not_org empty - skip');
                            continue;
                        }

                        notOrgArray = ParseJson(notOrgValue);
                        result.GetOptProperty('debug').push('ParseJson done: notOrgArray count=' + ArrayCount(notOrgArray));

                        for (item in notOrgArray) {
                            try {
                                orgId = OptInt(String(item.__value));
                                if (orgId != 0) {
                                    notOrgIds.push(orgId);
                                    result.GetOptProperty('debug').push('Excluded org ID: ' + orgId);
                                }
                            } catch(itemErr) {
                                result.GetOptProperty('debug').push('ERROR parsing not_org item: ' + String(itemErr));
                            }
                        }
                    }
                }
            } catch(docErr) {
                result.GetOptProperty('debug').push('ERROR loading not_org doc: ' + String(docErr));
            }
        }
    }

    result.GetOptProperty('debug').push('END: total notOrgIds count=' + ArrayCount(notOrgIds));
} catch(notOrgErr) {
    result.GetOptProperty('debug').push('ERROR loading not_org: ' + String(notOrgErr));
}
    try {
        result.GetOptProperty('debug').push('START: loading subNameMap');
        arrVars = tools.xquery("for $elem in custom_web_templates where $elem/code = 'testSd' return $elem");
        result.GetOptProperty('debug').push('xquery done: arrVars count = ' + ArrayCount(arrVars));
        if (ArrayCount(arrVars) == 0) {
            result.GetOptProperty('debug').push('PROBLEM: no custom_web_templates with code="testSd" found!');
        }
        for (elem in arrVars) {
            try {

                elemOpen = tools.open_doc(elem.id);

                if (elemOpen == undefined)
                    continue;
                elemTop = elemOpen.TopElem.wvars;

                if (elemTop == undefined)
                    continue;
                for (wvar in elemTop) {

                    if (wvar.name == 'vars1') {

                        vars1Value = String(wvar.value);
                        if (vars1Value == '') {
                            result.GetOptProperty('debug').push('vars1Value empty - skip');
                            continue;
                        }

                        vars1Array = ParseJson(vars1Value);

                        if (ArrayCount(vars1Array) == 0) {
                            result.GetOptProperty('debug').push('PROBLEM: vars1Array empty after parse!');
                        }
                        for (item in vars1Array) {
                            try {
                                subId = String(item.__value);
                                subName = String(item.comment);

                                // if (subId == '' || subName == '') continue;  // закомментил, как у тебя
                                subNameMap.push({
                                    id: OptInt(subId),
                                    name: subName
                                });

                            } catch (itemErr) {

                            }
                        }
                        result.GetOptProperty('debug').push('vars1 loop done, subNameMap now count=' + ArrayCount(subNameMap));
                    }
                    if (wvar.name == 'org_id') {
                        orgIdDop = String(wvar.value);
                        result.GetOptProperty('debug').push('org_id: "' + orgIdDop + '"');
                        if (orgIdDop == '')
                            continue;
                    }
                }
                result.GetOptProperty('debug').push('elem done for id=' + String(elem.id));
            } catch (docErr) {
                result.GetOptProperty('debug').push('ERROR open_doc/elemTop for id=' + String(elem.id) + ': ' + String(docErr));
            }
        }
        result.GetOptProperty('debug').push('END: total subNameMap count=' + ArrayCount(subNameMap));
    } catch (mapErr) {
        result.GetOptProperty('debug').push('OUTER ERROR: ' + String(mapErr));
    }

    function startsWithHiddenPrefix(name) {
        try {
            normalizedName = StrLowerCase(Trim(StrReplace(name, '"', '')));
            for (prefix in hiddenPrefixes) {
                normalizedPrefix = StrLowerCase(Trim(prefix));
                if (StrBegins(normalizedName, normalizedPrefix, false)) return true;
            }
        } catch(e) {}
        return false;
    }
	function getOrgSortOrder(org_id) {
    try {
        org_id_str = String(org_id);
        if (sortOrgMap.HasProperty(org_id_str)) {
            return sortOrgMap.GetOptProperty(org_id_str);
        }
    } catch(e) {}
    return 999999;
}
function sortOrganizations(org_array) {
    try {
        // Разделяем на две группы: с сортировкой и без
        sorted_orgs = [];
        unsorted_orgs = [];

        for (org in org_array) {
            org_id_str = String(org.GetOptProperty('id', ''));
            if (sortOrgMap.HasProperty(org_id_str)) {
                sort_order = sortOrgMap.GetOptProperty(org_id_str);
                org.SetProperty('_sort_order', sort_order);
                sorted_orgs.push(org);
                result.GetOptProperty('debug').push('Org ' + org.GetOptProperty('name', '') + ' has sort order: ' + sort_order);
            } else {
                unsorted_orgs.push(org);
            }
        }

        // Сортируем группу с заданным порядком
        sorted_orgs = ArraySort(sorted_orgs, '_sort_order', '+');

        // Объединяем: сначала отсортированные, потом остальные
        result_array = [];
        for (org in sorted_orgs) {
            result_array.push(org);
        }
        for (org in unsorted_orgs) {
            result_array.push(org);
        }

        return result_array;
    } catch(sortErr) {
        result.GetOptProperty('debug').push('ERROR in sortOrganizations: ' + String(sortErr));
        return org_array; // Возвращаем исходный массив в случае ошибки
    }
}
function sortSubdivisionsByName(sub_array, org_name) {
    try {
        // Если org_name не передан или нет карты - используем старую логику
        if (org_name == undefined || org_name == '' || !sortOrgSubdivisionsMap.HasProperty(org_name)) {
            result.GetOptProperty('debug').push('No custom sort for org "' + org_name + '", using default');
            return ArraySort(sub_array, 'name', '+');
        }

        org_sort_map = sortOrgSubdivisionsMap.GetOptProperty(org_name);

        sorted_subs = [];
        unsorted_subs = [];

        for (sub in sub_array) {
            sub_name = sub.GetOptProperty('name', '');
            sub_name_lower = StrLowerCase(Trim(sub_name));

            if (org_sort_map.HasProperty(sub_name_lower)) {
                sort_order = org_sort_map.GetOptProperty(sub_name_lower);
                sub.SetProperty('_sort_order', sort_order);
                sorted_subs.push(sub);
                result.GetOptProperty('debug').push('Org "' + org_name + '": "' + sub_name + '" = order ' + sort_order);
            } else {
                unsorted_subs.push(sub);
            }
        }

        sorted_subs = ArraySort(sorted_subs, '_sort_order', '+');
        unsorted_subs = ArraySort(unsorted_subs, 'name', '+');

        result_array = [];
        for (sub in sorted_subs) {
            result_array.push(sub);
        }
        for (sub in unsorted_subs) {
            result_array.push(sub);
        }

        return result_array;
    } catch(sortErr) {
        result.GetOptProperty('debug').push('ERROR in sortSubdivisionsByName: ' + String(sortErr));
        return sub_array;
    }
}
function isOrgExcluded(org_id) {
    try {
        org_id_int = OptInt(org_id);
        for (i = 0; i < ArrayCount(notOrgIds); i++) {
            if (notOrgIds[i] == org_id_int) {
                return true;
            }
        }
    } catch(e) {}
    return false;
}
function getSubNameById(subId) {
    try {
        for (sub in subNameMap) {
            if (OptInt(subId) == sub.id) return sub.name;  // <-- OptInt(subId) для конверта hex-string в int
        }
    } catch(e) {}
    return null;
}

function normalizeSubdivisionName(name) {
    try {
        normalized = StrReplace(name, '"', '');

        // Убираем ВСЕ точки
        while (StrContains(normalized, '.', false)) {
            normalized = StrReplace(normalized, '.', '');
        }

        // Убираем другие символы-разделители
        normalized = StrReplace(normalized, '-', '');
        normalized = StrReplace(normalized, '_', '');
        normalized = StrReplace(normalized, '/', '');
        normalized = StrReplace(normalized, '\\', '');

        // Убираем множественные пробелы
        normalized = Trim(normalized);
        while (StrContains(normalized, '  ', false)) {
            normalized = StrReplace(normalized, '  ', ' ');
        }

        // Приводим к lowercase
        normalized = StrLowerCase(normalized);

        result.GetOptProperty('debug').push('Normalized: "' + name + '" -> "' + normalized + '"');

        return normalized;
    } catch(e) {
        return StrLowerCase(name);
    }
}

function splitByComma(str) {
    result_arr = [];
    try {
        result.GetOptProperty('debug').push('SplitByComma input len=' + StrLen(str) + ', raw=' + str);
        current = '';
        for (i = 0; i < StrLen(str); i++) {
            char = StrLeftRange(StrRightRangePos(str, i), 1);
            if (char == ',') {
                if (current != '') {
                    trimmed = Trim(current);
                    result_arr.push(trimmed);
                    result.GetOptProperty('debug').push('Split part: ' + trimmed + ' (len=' + StrLen(trimmed) + ')');
                }
                current = '';
            } else {
                current = current + char;
            }
        }
        if (current != '') {
            trimmed = Trim(current);
            result_arr.push(trimmed);
            result.GetOptProperty('debug').push('Split last part: ' + trimmed + ' (len=' + StrLen(trimmed) + ')');
        }
        result.GetOptProperty('debug').push('Split result count=' + ArrayCount(result_arr));
    } catch(e) {
        result.GetOptProperty('debug').push('Split error: ' + String(e));
    }
    return result_arr;
}

    function hasCollaboratorsRecursive(sub_id, depth) {
        try {
            if (depth == undefined) depth = 0;
            if (depth > 10) return false;
            direct_count_query = "for $elem in collaborators where $elem/position_parent_id = " + sub_id + " and $elem/is_dismiss = false() return $elem";
            direct_count_result = ArraySelectAll(tools.xquery(direct_count_query));
            if (ArrayCount(direct_count_result) > 0) return true;
            children_query = "for $elem in subdivisions where $elem/parent_object_id = " + sub_id + " and $elem/is_disbanded != true() return $elem";
            children_list = ArraySelectAll(tools.xquery(children_query));
            for (child in children_list) {
                if (hasCollaboratorsRecursive(String(child.id), depth + 1)) return true;
            }
        } catch(e) {}
        return false;
    }

    function hasSubdivisionsWithCollaborators(sub_id) {
    try {
        children_query = "for $elem in subdivisions where $elem/parent_object_id = " + sub_id + " and $elem/is_disbanded != true() return $elem";
        children_list = ArraySelectAll(tools.xquery(children_query));
        for (child in children_list) {
            if (hasCollaboratorsRecursive(String(child.id), 0)) return true;
        }
        return false;
    } catch(e) {
        return false;
    }
}

    function checkBirthday(birthDate) {
        if (birthDate == null || birthDate == '') return false;
        try {
            todayDate = Date();
            userDateStr = StrDate(birthDate, false);
            todayDateStr = StrDate(todayDate, false);
            userDayMonth = StrLeftCharRange(userDateStr, 5);
            todayDayMonth = StrLeftCharRange(todayDateStr, 5);
            return userDayMonth == todayDayMonth;
        } catch(e) {
            return false;
        }
    }
function getFuncManagerForSubdivision(sub_id, only_subdivision_catalog) {
    if (only_subdivision_catalog == undefined) only_subdivision_catalog = false;

    func_manager_query_subdivision = "for $elem in func_managers " +
                        "where $elem/object_id = " + sub_id +
                        " and $elem/catalog = 'subdivision' and $elem/boss_type_id = 6148914691236517290" +
                        " and $elem/person_id != null() and $elem/person_id != '' " +
                        " and ($elem/is_finished = false() or $elem/is_finished = null()) " +
                        " return $elem";
    func_manager_result_subdivision = ArraySelectAll(tools.xquery(func_manager_query_subdivision));

    if (ArrayCount(func_manager_result_subdivision) > 0 && func_manager_result_subdivision[0].person_id != null) {
        return String(func_manager_result_subdivision[0].person_id);
    }

    // Если only_subdivision_catalog = true, то НЕ проверяем position
    if (only_subdivision_catalog) {
        return '';
    }

    func_manager_query_position = "for $elem in func_managers " +
                       "where $elem/parent_id = " + sub_id +
                       " and $elem/catalog = 'position' " +
                       " and $elem/is_native = true() " +
                       " and $elem/person_id != null() and $elem/person_id != '' " +
                       " and ($elem/is_finished = false() or $elem/is_finished = null()) " +
                       " return $elem";
    func_manager_result_position = ArraySelectAll(tools.xquery(func_manager_query_position));

    if (ArrayCount(func_manager_result_position) > 0 && func_manager_result_position[0].person_id != null) {
        return String(func_manager_result_position[0].person_id);
    }

    return '';
}

function isPersonFuncManager(person_id) {
    try {
        // Для position проверяем is_native=true
        // Для subdivision просто наличие записи (независимо от object_id/parent_id)
        fm_check_query = "for $elem in func_managers " +
                        "where $elem/person_id = " + person_id +
                        " and (($elem/catalog = 'position' and $elem/is_native = true()) or $elem/catalog = 'subdivision') " +
                        " and ($elem/is_finished = false() or $elem/is_finished = null()) " +
                        " return $elem";
        fm_check_result = ArraySelectAll(tools.xquery(fm_check_query));
        return ArrayCount(fm_check_result) > 0;
    } catch(e) {}
    return false;
}

// === КОНЕЦ ПЕРВОЙ ЧАСТИ (ФУНКЦИИ-УТИЛИТЫ) ===
%>
