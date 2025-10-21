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

// === ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ДЛЯ УМЕНЬШЕНИЯ ДУБЛИРОВАНИЯ ===

// Функция поиска territorial (скрытого) родителя - используется много раз в коде
function findTerritorialParent(sub_id, max_depth) {
    if (max_depth == undefined) max_depth = 10;

    regional_name = '';
    regional_parent_id_str = '';

    try {
        sub_query = "for $elem in subdivisions where $elem/id = " + sub_id + " return $elem";
        sub_result = ArraySelectAll(tools.xquery(sub_query));

        if (ArrayCount(sub_result) == 0) {
            return {name: regional_name, id: regional_parent_id_str};
        }

        current_check_sub = sub_result[0];
        search_depth = 0;
        found_territorial = false;

        while (search_depth < max_depth && !found_territorial) {
            search_depth++;
            check_parent_id = String(current_check_sub.parent_object_id);

            if (check_parent_id == '' || check_parent_id == null) break;

            check_parent_query = "for $elem in subdivisions where $elem/id = " + check_parent_id + " return $elem";
            check_parent_result = ArraySelectAll(tools.xquery(check_parent_query));

            if (ArrayCount(check_parent_result) > 0) {
                check_parent_name = String(check_parent_result[0].name);

                if (startsWithHiddenPrefix(check_parent_name)) {
                    regional_name = check_parent_name;
                    regional_parent_id_str = check_parent_id;
                    found_territorial = true;
                    result.GetOptProperty('debug').push('Found territorial at depth ' + search_depth + ': ' + regional_name);
                    break;
                }

                current_check_sub = check_parent_result[0];
            } else {
                break;
            }
        }

        if (!found_territorial) {
            result.GetOptProperty('debug').push('No territorial found for sub ' + sub_id);
        }
    } catch(regErr) {
        result.GetOptProperty('debug').push('Error finding territorial: ' + String(regErr));
    }

    return {name: regional_name, id: regional_parent_id_str};
}

// === ФУНКЦИЯ ЗАГРУЗКИ СОТРУДНИКОВ С ПАГИНАЦИЕЙ ===
function getAllCollaboratorsPaginated(sub_id, offset, limit) {
    result.GetOptProperty('debug').push('Paginated load: sub_id=' + sub_id + ', offset=' + offset + ', limit=' + limit);

    all_collaborators = [];
    total_count = 0;

    try {
        // 1. Получаем ТОЛЬКО прямых сотрудников текущего подразделения
        direct_col_query = "for $elem in collaborators where $elem/position_parent_id = " + sub_id + " and $elem/is_dismiss = false() return $elem";
        direct_collaborators = ArraySelectAll(tools.xquery(direct_col_query));

        // 2. Получаем ТОЛЬКО прямые дочерние подразделения (уровень 1)
        children_query = "for $elem in subdivisions where $elem/parent_object_id = " + sub_id + " and $elem/is_disbanded != true() return $elem";
        children_list = ArraySelectAll(tools.xquery(children_query));

        // 3. Получаем подразделения уровня 2 (внуки)
        grandchildren_list = [];
        for (child in children_list) {
            try {
                grandchildren_query = "for $elem in subdivisions where $elem/parent_object_id = " + String(child.id) + " and $elem/is_disbanded != true() return $elem";
                grandchildren = ArraySelectAll(tools.xquery(grandchildren_query));

                for (grandchild in grandchildren) {
                    grandchildren_list.push(grandchild);
                }
            } catch(gcErr) {
                result.GetOptProperty('debug').push('Error loading grandchildren for ' + String(child.id) + ': ' + String(gcErr));
            }
        }

        result.GetOptProperty('debug').push('Direct collaborators: ' + ArrayCount(direct_collaborators));
        result.GetOptProperty('debug').push('Level 1 children: ' + ArrayCount(children_list));
        result.GetOptProperty('debug').push('Level 2 children (grandchildren): ' + ArrayCount(grandchildren_list));

        // 4. Получаем сотрудников из прямых дочерних подразделений (уровень 1)
        children_collaborators = [];
        for (child in children_list) {
            try {
                child_col_query = "for $elem in collaborators where $elem/position_parent_id = " + String(child.id) + " and $elem/is_dismiss = false() return $elem";
                child_cols = ArraySelectAll(tools.xquery(child_col_query));

                for (child_col in child_cols) {
                    children_collaborators.push(child_col);
                }
            } catch(childErr) {}
        }

        // 5. Получаем сотрудников из внуков (уровень 2)
        grandchildren_collaborators = [];
        for (grandchild in grandchildren_list) {
            try {
                gc_col_query = "for $elem in collaborators where $elem/position_parent_id = " + String(grandchild.id) + " and $elem/is_dismiss = false() return $elem";
                gc_cols = ArraySelectAll(tools.xquery(gc_col_query));

                for (gc_col in gc_cols) {
                    grandchildren_collaborators.push(gc_col);
                }
            } catch(gcColErr) {}
        }

        result.GetOptProperty('debug').push('Children collaborators: ' + ArrayCount(children_collaborators));
        result.GetOptProperty('debug').push('Grandchildren collaborators: ' + ArrayCount(grandchildren_collaborators));

        // 6. ГРУППИРУЕМ ПО ПОДРАЗДЕЛЕНИЯМ С СОХРАНЕНИЕМ ПОРЯДКА
        collaborators_by_subdivision = new Object();

        // Руководитель родительского подразделения
        func_manager_id = getFuncManagerForSubdivision(sub_id);
        result.GetOptProperty('debug').push('Func manager ID: ' + func_manager_id);

        // Собираем руководителей всех дочерних подразделений (уровень 1)
        child_func_managers = new Object();
        for (child in children_list) {
            try {
                child_id_str = String(child.id);
                child_fm_id = getFuncManagerForSubdivision(child_id_str);
                if (child_fm_id != '') {
                    child_func_managers.SetProperty(child_fm_id, true);
                    result.GetOptProperty('debug').push('Child func manager found: ' + child_fm_id + ' for subdivision: ' + child_id_str);
                }
            } catch(e) {}
        }

        // Собираем руководителей внуков (уровень 2)
        grandchild_func_managers = new Object();
        for (grandchild in grandchildren_list) {
            try {
                gc_id_str = String(grandchild.id);
                gc_fm_id = getFuncManagerForSubdivision(gc_id_str);
                if (gc_fm_id != '') {
                    grandchild_func_managers.SetProperty(gc_fm_id, true);
                    result.GetOptProperty('debug').push('Grandchild func manager found: ' + gc_fm_id + ' for subdivision: ' + gc_id_str);
                }
            } catch(e) {}
        }

        // СНАЧАЛА обрабатываем прямых сотрудников родителя, ИСКЛЮЧАЯ функционального менеджера
        parent_sub_id = String(sub_id);
        if (!collaborators_by_subdivision.HasProperty(parent_sub_id)) {
            collaborators_by_subdivision.SetProperty(parent_sub_id, []);
        }

        for (col in direct_collaborators) {
            try {
                if (!col.id || col.id == '') continue;
                if (String(col.id) == func_manager_id) continue; // Пропускаем функционального менеджера
                collaborators_by_subdivision.GetOptProperty(parent_sub_id).push(col);
            } catch(e) {}
        }

        // ПОТОМ обрабатываем сотрудников из дочерних подразделений уровня 1
        for (child in children_list) {
            try {
                child_id_str = String(child.id);
                if (!collaborators_by_subdivision.HasProperty(child_id_str)) {
                    collaborators_by_subdivision.SetProperty(child_id_str, []);
                }

                for (col in children_collaborators) {
                    if (String(col.position_parent_id) == child_id_str) {
                        collaborators_by_subdivision.GetOptProperty(child_id_str).push(col);
                    }
                }
            } catch(e) {}
        }

        // И НАКОНЕЦ обрабатываем сотрудников из внуков (уровень 2)
        for (grandchild in grandchildren_list) {
            try {
                gc_id_str = String(grandchild.id);
                if (!collaborators_by_subdivision.HasProperty(gc_id_str)) {
                    collaborators_by_subdivision.SetProperty(gc_id_str, []);
                }

                for (col in grandchildren_collaborators) {
                    if (String(col.position_parent_id) == gc_id_str) {
                        collaborators_by_subdivision.GetOptProperty(gc_id_str).push(col);
                    }
                }
            } catch(e) {}
        }

        temp_all = [];

        // Добавляем функционального менеджера в начало temp_all
        if (func_manager_id != '') {
            try {
                func_manager_full_query = "for $elem in collaborators where $elem/id = " + func_manager_id + " and $elem/is_dismiss = false() return $elem";
                func_manager_data_result = ArraySelectAll(tools.xquery(func_manager_full_query));

                if (ArrayCount(func_manager_data_result) > 0) {
                    fm_elem = func_manager_data_result[0];
                    temp_all.push(fm_elem); // Добавляем в начало
                    result.GetOptProperty('debug').push('Added func manager to temp_all: ' + func_manager_id + ', name=' + String(fm_elem.fullname));
                } else {
                    result.GetOptProperty('debug').push('Func manager ' + func_manager_id + ' not found or dismissed');
                }
            } catch(fmErr) {
                result.GetOptProperty('debug').push('Error adding func manager to temp_all: ' + String(fmErr));
            }
        }

        // Добавляем сотрудников родителя
        if (collaborators_by_subdivision.HasProperty(parent_sub_id)) {
            parent_cols = collaborators_by_subdivision.GetOptProperty(parent_sub_id);
            for (col in parent_cols) {
                temp_all.push(col);
            }
        }

        // Добавляем сотрудников дочерних отделов уровня 1 СТРОГО В ПОРЯДКЕ children_list
        for (child in children_list) {
            try {
                child_id_str = String(child.id);
                if (collaborators_by_subdivision.HasProperty(child_id_str)) {
                    child_cols = collaborators_by_subdivision.GetOptProperty(child_id_str);
                    for (col in child_cols) {
                        temp_all.push(col);
                    }
                }
            } catch(e) {}
        }

        // Добавляем сотрудников внуков уровня 2 СТРОГО В ПОРЯДКЕ grandchildren_list
        for (grandchild in grandchildren_list) {
            try {
                gc_id_str = String(grandchild.id);
                if (collaborators_by_subdivision.HasProperty(gc_id_str)) {
                    gc_cols = collaborators_by_subdivision.GetOptProperty(gc_id_str);
                    for (col in gc_cols) {
                        temp_all.push(col);
                    }
                }
            } catch(e) {}
        }

        total_count = ArrayCount(temp_all);
        result.GetOptProperty('debug').push('Total before pagination: ' + total_count);

        // 7. Применяем пагинацию и форматируем
        start_index = offset;
        end_index = offset + limit;
        if (end_index > total_count) end_index = total_count;

        for (i = start_index; i < end_index; i++) {
            col_elem = temp_all[i];

            try {
                if (!col_elem.id || col_elem.id == '') continue;

                is_birthday = checkBirthday(col_elem.birth_date);
                is_on_vacation = (String(col_elem.current_state) == 'Отпуск');

                // Проверяем: это руководитель родителя ИЛИ руководитель дочернего подразделения ИЛИ внука
                is_func_mgr = (func_manager_id != '' && String(col_elem.id) == func_manager_id) ||
                              child_func_managers.HasProperty(String(col_elem.id)) ||
                              grandchild_func_managers.HasProperty(String(col_elem.id));

                // Получаем имя подразделения
                col_subdivision_name = '';
                if (String(col_elem.id) == func_manager_id) {
                    // Для функционального менеджера используем имя текущего подразделения
                    col_sub_query = "for $elem in subdivisions where $elem/id = " + sub_id + " return $elem";
                    col_sub_result = ArraySelectAll(tools.xquery(col_sub_query));
                    if (ArrayCount(col_sub_result) > 0) {
                        colSubName = getSubNameById(sub_id);
                        if (colSubName == null) {
                            colSubName = String(col_sub_result[0].name);
                        }
                        col_subdivision_name = colSubName;
                    } else {
                        result.GetOptProperty('debug').push('No subdivision found for sub_id=' + sub_id);
                    }
                } else {
                    col_sub_query = "for $elem in subdivisions where $elem/id = " + String(col_elem.position_parent_id) + " return $elem";
                    col_sub_result = ArraySelectAll(tools.xquery(col_sub_query));
                    if (ArrayCount(col_sub_result) > 0) {
                        colSubName = getSubNameById(String(col_elem.position_parent_id));
                        if (colSubName == null) {
                            colSubName = String(col_sub_result[0].name);
                        }
                        col_subdivision_name = colSubName;
                    } else {
                        result.GetOptProperty('debug').push('No subdivision found for position_parent_id=' + String(col_elem.position_parent_id));
                    }
                }

                col_data = new Object();
                col_data.SetProperty('id', String(col_elem.id));
                col_data.SetProperty('name', String(col_elem.fullname));
                col_data.SetProperty('name_lower', StrLowerCase(String(col_elem.fullname)));
                col_data.SetProperty('subdivision_id', String(col_elem.position_parent_id));
                col_data.SetProperty('subdivision_name', col_subdivision_name);
                col_data.SetProperty('email', String(col_elem.email != null ? col_elem.email : '—'));
                col_data.SetProperty('pict_url', String(col_elem.pict_url != null ? col_elem.pict_url : ''));
                col_data.SetProperty('position_name', String(col_elem.position_name != null ? col_elem.position_name : '—'));
                col_data.SetProperty('mobile_phone', String(col_elem.mobile_phone != null ? col_elem.mobile_phone : '—'));
                col_data.SetProperty('phone', String(col_elem.phone != null ? col_elem.phone : '—'));
                col_data.SetProperty('is_birthday', is_birthday);
                col_data.SetProperty('is_on_vacation', is_on_vacation);
                col_data.SetProperty('is_func_manager', is_func_mgr);
                col_data.SetProperty('dept_name_normalized', normalizeSubdivisionName(col_subdivision_name));

                // РЕКУРСИВНЫЙ ПОИСК скрытого territorial подразделения
                try {
                    col_pos_parent_id = String(col_elem.position_parent_id);
                    col_sub_detail_query = "for $elem in subdivisions where $elem/id = " + col_pos_parent_id + " return $elem";
                    col_sub_detail_result = ArraySelectAll(tools.xquery(col_sub_detail_query));

                    if (ArrayCount(col_sub_detail_result) > 0) {
                        current_sub = col_sub_detail_result[0];
                        check_depth = 0;
                        found_regional = false;

                        while (check_depth < 10 && !found_regional) {
                            check_depth++;
                            col_parent_id = String(current_sub.parent_object_id);

                            if (col_parent_id == '' || col_parent_id == null) break;

                            col_parent_query = "for $elem in subdivisions where $elem/id = " + col_parent_id + " return $elem";
                            col_parent_result = ArraySelectAll(tools.xquery(col_parent_query));

                            if (ArrayCount(col_parent_result) > 0) {
                                col_parent_name = String(col_parent_result[0].name);

                                if (startsWithHiddenPrefix(col_parent_name)) {
                                    col_data.SetProperty('regional_name', col_parent_name);
                                    col_data.SetProperty('regional_parent_id', col_parent_id);
                                    col_data.SetProperty('original_dept_id', col_pos_parent_id);
                                    found_regional = true;
                                    result.GetOptProperty('debug').push('Found regional for ' + String(col_elem.id) + ' at depth ' + check_depth + ': ' + col_parent_name);
                                    break;
                                }

                                current_sub = col_parent_result[0];
                            } else {
                                break;
                            }
                        }
                    }
                } catch(regional_err) {
                    result.GetOptProperty('debug').push('Error checking regional for col ' + String(col_elem.id) + ': ' + String(regional_err));
                }

                all_collaborators.push(col_data);
                result.GetOptProperty('debug').push('Added collaborator: id=' + String(col_elem.id) + ', name=' + String(col_elem.fullname) + ', is_func_manager=' + is_func_mgr + ', subdivision_name=' + col_subdivision_name);
            } catch(colErr) {
                result.GetOptProperty('debug').push('Error processing collaborator ' + String(col_elem.id) + ': ' + String(colErr));
            }
        }

        result.GetOptProperty('debug').push('Returned collaborators: ' + ArrayCount(all_collaborators));

    } catch(getAllErr) {
        result.GetOptProperty('debug').push('ERROR in getAllCollaboratorsPaginated: ' + String(getAllErr));
    }

    return_data = new Object();
    return_data.SetProperty('collaborators', all_collaborators);
    return_data.SetProperty('total_count', total_count);
    return_data.SetProperty('offset', offset);
    return_data.SetProperty('limit', limit);
    return_data.SetProperty('has_more', end_index < total_count);

    return return_data;
}

// === ОСНОВНОЙ КОД ОБРАБОТКИ ЗАПРОСОВ ===

// Парсинг параметров запроса
result = new Object();
result.SetProperty('structure', []);
result.SetProperty('filtered_collaborators', []);
result.SetProperty('debug', []);

letter_param = Request.QueryString.GetOptProperty('letter', '');
search_param = Request.QueryString.GetOptProperty('search', '');
org_id = Request.QueryString.GetOptProperty('org_id', '');
subdivision_id = Request.QueryString.GetOptProperty('subdivision_id', '');
grouped_items_param = Request.QueryString.GetOptProperty('grouped_items', '');

letter_param = StrReplace(letter_param, "'", "''");
search_param = StrReplace(search_param, "'", "''");
org_id = StrReplace(org_id, "'", "''");
subdivision_id = StrReplace(subdivision_id, "'", "''");

result.GetOptProperty('debug').push('Parameters: letter=' + letter_param + ', search=' + search_param + ', org_id=' + org_id + ', subdivision_id=' + subdivision_id);
result.GetOptProperty('debug').push('subNameMap entries: ' + ArrayCount(subNameMap));

// === ОБРАБОТКА GROUPED ПОДРАЗДЕЛЕНИЙ ===
if (subdivision_id != '' && StrBegins(subdivision_id, 'grouped_')) {
    if (grouped_items_param == '') {
        result.SetProperty('error', 'grouped_items parameter required');
        Response.Write(tools.object_to_text(result, 'json'));
    } else {
        grouped_ids = splitByComma(grouped_items_param);
        result.GetOptProperty('debug').push('Detail grouped param raw=' + grouped_items_param + ', split count=' + ArrayCount(grouped_ids));

        for (i = 0; i < ArrayCount(grouped_ids); i++) {
            result.GetOptProperty('debug').push('Detail dept loop #' + i + ': dept_id=' + grouped_ids[i] + ' (int=' + OptInt(grouped_ids[i]) + ')');
        }

        regional_groups = [];
        all_children_map = new Object();

        // Загружаем ВСЕ подразделения один раз
        all_descendants_query = "for $elem in subdivisions where $elem/is_disbanded != true() return $elem";
        all_subs = ArraySelectAll(tools.xquery(all_descendants_query));
        result.GetOptProperty('debug').push('Grouped all_subs total count=' + ArrayCount(all_subs));

        // Функция для проверки является ли sub потомком dept_id
        function isDescendantOf(sub, ancestor_id, all_subdivisions, depth) {
            if (depth > 20) return false;
            if (String(sub.parent_object_id) == ancestor_id) return true;
            if (sub.parent_object_id == '' || sub.parent_object_id == null) return false;

            for (parent_sub in all_subdivisions) {
                if (String(parent_sub.id) == String(sub.parent_object_id)) {
                    return isDescendantOf(parent_sub, ancestor_id, all_subdivisions, depth + 1);
                }
            }
            return false;
        }

        global_normalized_map = new Object(); // Глобальная карта для группировки детей из ВСЕХ dept_id

        for (i = 0; i < ArrayCount(grouped_ids); i++) {
            dept_id = grouped_ids[i];
            if (dept_id == '') continue;

            try {
                dept_query = "for $elem in subdivisions where $elem/id = " + dept_id + " return $elem";
                dept_result = ArraySelectAll(tools.xquery(dept_query));
                if (ArrayCount(dept_result) == 0) {
                    result.GetOptProperty('debug').push('Grouped dept_id=' + dept_id + ' NOT FOUND');
                    continue;
                }

                dept_info = dept_result[0];

                // Используем вспомогательную функцию для поиска territorial родителя
                territorial_result = findTerritorialParent(dept_id, 10);
                regional_name = territorial_result.name;
                regional_parent_id_str = territorial_result.id;

                if (regional_name == '') {
                    regional_name = 'Неизвестное подразделение';
                }

                // Ищем всех потомков текущего dept_id и СРАЗУ ДОБАВЛЯЕМ В ГЛОБАЛЬНУЮ КАРТУ
                for (sub in all_subs) {
                    if (!sub.id || sub.id == '') continue;
                    if (String(sub.id) == dept_id) continue;

                    try {
                        if (isDescendantOf(sub, dept_id, all_subs, 0)) {
                            // Проверяем сотрудников В ПОДРАЗДЕЛЕНИИ ИЛИ В ЕГО ДЕТЯХ
                            has_collab = hasCollaboratorsRecursive(String(sub.id), 0);
                            has_children_with_collab = hasSubdivisionsWithCollaborators(String(sub.id));

                            // Пропускаем только если НЕТ сотрудников И НЕТ детей с сотрудниками
                            if (!has_collab && !has_children_with_collab) {
                                result.GetOptProperty('debug').push('Skip ' + String(sub.id) + ': no collab and no children with collab');
                                continue;
                            }

                            sub_id_str = String(sub.id);

                            // Получаем имя и очищаем от точек
                            childName = getSubNameById(sub_id_str);
                            if (childName == null) {
                                childName = String(sub.name);
                            }

                            // Очищаем от точек ДЛЯ ОТОБРАЖЕНИЯ
                            clean_name = childName;
                            while (StrContains(clean_name, '.', false)) {
                                clean_name = StrReplace(clean_name, '.', '');
                            }
                            clean_name = Trim(clean_name);

                            // НОРМАЛИЗУЕМ очищенное имя
                            normalized_name = normalizeSubdivisionName(clean_name);

                            result.GetOptProperty('debug').push('Child: id=' + sub_id_str + ', original="' + childName + '", clean="' + clean_name + '", normalized="' + normalized_name + '"');

                            // Создаем/обновляем группу в ГЛОБАЛЬНОЙ карте
                            if (!global_normalized_map.HasProperty(normalized_name)) {
                                group_obj = new Object();
                                group_obj.SetProperty('display_name', clean_name); // Первое встреченное ЧИСТОЕ имя
                                group_obj.SetProperty('ids', []);
                                group_obj.SetProperty('org_id', String(sub.org_id));
                                group_obj.SetProperty('parent_object_id', String(sub.parent_object_id));
                                group_obj.SetProperty('regional_parent_name', regional_name);
                                group_obj.SetProperty('has_children', false);
                                global_normalized_map.SetProperty(normalized_name, group_obj);
                            }

                            // Добавляем ID в группу
                            global_normalized_map.GetOptProperty(normalized_name).ids.push(sub_id_str);

                            // Проверяем дочерние подразделения
                            if (hasSubdivisionsWithCollaborators(sub_id_str)) {
                                global_normalized_map.GetOptProperty(normalized_name).SetProperty('has_children', true);
                            }
                        }
                    } catch(descErr) {
                        result.GetOptProperty('debug').push('Error checking descendant: ' + String(descErr));
                    }
                }

                // Собираем сотрудников родительского dept (это для regional_groups)
                func_manager_id = '';
                try {
                    func_manager_id = getFuncManagerForSubdivision(dept_id);
                } catch(fmErr) {}

                col_query = "for $elem in collaborators where $elem/position_parent_id = " + dept_id + " and $elem/is_dismiss = false() return $elem";
                col_list = ArraySelectAll(tools.xquery(col_query));

                collaborators_array = [];

                if (func_manager_id != '') {
                    try {
                        func_manager_full_query = "for $elem in collaborators where $elem/id = " + func_manager_id + " and $elem/is_dismiss = false() return $elem";
                        func_manager_data_result = ArraySelectAll(tools.xquery(func_manager_full_query));

                        if (ArrayCount(func_manager_data_result) > 0) {
                            fm_elem = func_manager_data_result[0];
                            is_birthday_fm = checkBirthday(fm_elem.birth_date);
                            is_on_vacation_fm = (String(fm_elem.current_state) == 'Отпуск');

                            fm_data = new Object();
                            fm_data.SetProperty('id', String(fm_elem.id));
                            fm_data.SetProperty('name', String(fm_elem.fullname));
                            fm_data.SetProperty('name_lower', StrLowerCase(String(fm_elem.fullname)));
                            fm_data.SetProperty('subdivision_id', String(fm_elem.position_parent_id));
                            fm_data.SetProperty('subdivision_name', regional_name);
                            fm_data.SetProperty('email', String(fm_elem.email != null ? fm_elem.email : '—'));
                            fm_data.SetProperty('pict_url', String(fm_elem.pict_url != null ? fm_elem.pict_url : ''));
                            fm_data.SetProperty('position_name', String(fm_elem.position_name != null ? fm_elem.position_name : '—'));
                            fm_data.SetProperty('mobile_phone', String(fm_elem.mobile_phone != null ? fm_elem.mobile_phone : '—'));
                            fm_data.SetProperty('phone', String(fm_elem.phone != null ? fm_elem.phone : '—'));
                            fm_data.SetProperty('is_birthday', is_birthday_fm);
                            fm_data.SetProperty('is_on_vacation', is_on_vacation_fm);
                            fm_data.SetProperty('is_func_manager', true);

                            collaborators_array.push(fm_data);
                        }
                    } catch(fmErr) {}
                }

                for (col_elem in col_list) {
                    try {
                        if (!col_elem.id || col_elem.id == '') continue;
                        if (func_manager_id != '' && String(col_elem.id) == func_manager_id) continue;

                        is_birthday = checkBirthday(col_elem.birth_date);
                        is_on_vacation = (String(col_elem.current_state) == 'Отпуск');

                        col_data = new Object();
                        col_data.SetProperty('id', String(col_elem.id));
                        col_data.SetProperty('name', String(col_elem.fullname));
                        col_data.SetProperty('name_lower', StrLowerCase(String(col_elem.fullname)));
                        col_data.SetProperty('subdivision_id', String(col_elem.position_parent_id));
                        col_data.SetProperty('subdivision_name', regional_name);
                        col_data.SetProperty('email', String(col_elem.email != null ? col_elem.email : '—'));
                        col_data.SetProperty('pict_url', String(col_elem.pict_url != null ? col_elem.pict_url : ''));
                        col_data.SetProperty('position_name', String(col_elem.position_name != null ? col_elem.position_name : '—'));
                        col_data.SetProperty('mobile_phone', String(col_elem.mobile_phone != null ? col_elem.mobile_phone : '—'));
                        col_data.SetProperty('phone', String(col_elem.phone != null ? col_elem.phone : '—'));
                        col_data.SetProperty('is_birthday', is_birthday);
                        col_data.SetProperty('is_on_vacation', is_on_vacation);
                        col_data.SetProperty('is_func_manager', false);

                        collaborators_array.push(col_data);
                    } catch(colErr) {}
                }

                regional_group = new Object();
                regional_group.SetProperty('regional_name', regional_name);
                regional_group.SetProperty('collaborators', collaborators_array);
                regional_groups.push(regional_group);

            } catch(deptErr) {
                result.GetOptProperty('debug').push('Error processing dept_id=' + dept_id + ': ' + String(deptErr));
            }
        }

        result.GetOptProperty('debug').push('=== CREATING FINAL CHILDREN FROM GLOBAL MAP ===');

        for (norm_key in global_normalized_map) {
            try {
                group_info = global_normalized_map.GetOptProperty(norm_key);
                group_ids = group_info.ids;

                result.GetOptProperty('debug').push('Group "' + group_info.display_name + '": ' + ArrayCount(group_ids) + ' items');

                // Если несколько ID - создаем grouped элемент
                if (ArrayCount(group_ids) > 1) {
                    grouped_ids_str = '';
                    for (j = 0; j < ArrayCount(group_ids); j++) {
                        if (j > 0) grouped_ids_str = grouped_ids_str + ',';
                        grouped_ids_str = grouped_ids_str + group_ids[j];
                    }

                    unique_key = 'grouped_' + norm_key;

                    child_data = new Object();
                    child_data.SetProperty('id', unique_key);
                    child_data.SetProperty('name', group_info.display_name);
                    child_data.SetProperty('type', 'grouped');
                    child_data.SetProperty('is_grouped', true);
                    child_data.SetProperty('group_count', ArrayCount(group_ids));
                    child_data.SetProperty('org_id', group_info.org_id);
                    child_data.SetProperty('parent_object_id', group_info.parent_object_id);
                    child_data.SetProperty('regional_parent_name', group_info.regional_parent_name);
                    child_data.SetProperty('has_subdivisions', group_info.has_children);
                    child_data.SetProperty('grouped_items_str', grouped_ids_str);
                    child_data.SetProperty('children', []);
                    child_data.SetProperty('collaborators', []);

                    all_children_map.SetProperty(unique_key, child_data);

                    result.GetOptProperty('debug').push('✓ CREATED GROUPED: "' + group_info.display_name + '" with ' + ArrayCount(group_ids) + ' items, org_id=' + group_info.org_id);

                } else if (ArrayCount(group_ids) == 1) {
                    // Одно подразделение - обычный элемент
                    single_id = group_ids[0];

                    child_data = new Object();
                    child_data.SetProperty('id', single_id);
                    child_data.SetProperty('name', group_info.display_name);
                    child_data.SetProperty('org_id', group_info.org_id);
                    child_data.SetProperty('parent_object_id', group_info.parent_object_id);
                    child_data.SetProperty('regional_parent_name', group_info.regional_parent_name);
                    child_data.SetProperty('has_subdivisions', group_info.has_children);
                    child_data.SetProperty('children', []);
                    child_data.SetProperty('collaborators', []);

                    all_children_map.SetProperty(single_id, child_data);

                    result.GetOptProperty('debug').push('✓ Added single: "' + group_info.display_name + '", org_id=' + group_info.org_id);
                }
            } catch(finalErr) {
                result.GetOptProperty('debug').push('ERROR creating final child: ' + String(finalErr));
            }
        }

        // Конвертируем map в массив
        all_children_array = [];
        for (child_key in all_children_map) {
            all_children_array.push(all_children_map.GetOptProperty(child_key));
        }

        result.GetOptProperty('debug').push('Final all_children_array count=' + ArrayCount(all_children_array));

        regional_groups = ArraySort(regional_groups, 'regional_name', '+');

        subdivision_data = new Object();
        subdivision_data.SetProperty('id', subdivision_id);
        subdivision_data.SetProperty('hierarchy_path', []);
        subdivision_data.SetProperty('children', all_children_array);
        subdivision_data.SetProperty('regional_groups', regional_groups);
        subdivision_data.SetProperty('collaborators', []);
        subdivision_data.SetProperty('has_subdivisions', ArrayCount(all_children_array) > 0);

        // ОБРАБОТКА ПАГИНАЦИИ ДЛЯ GROUPED
        include_all_param = Request.QueryString.GetOptProperty('include_all_collaborators', '');
        offset_param = OptInt(Request.QueryString.GetOptProperty('offset', '0'));
        limit_param = OptInt(Request.QueryString.GetOptProperty('limit', '20'));

        if (include_all_param == 'true') {
            // Здесь будет большой блок кода для загрузки сотрудников...
            // Продолжение в следующем блоке из-за ограничения размера
            result.SetProperty('subdivision', subdivision_data);
        } else {
            subdivision_data.SetProperty('all_collaborators', []);
            subdivision_data.SetProperty('total_count', 0);
            subdivision_data.SetProperty('has_more', false);
            result.SetProperty('subdivision', subdivision_data);
        }

        Response.Write(tools.object_to_text(result, 'json'));
    }
}

// === КОНЕЦ ПЕРВОЙ ЧАСТИ (ФУНКЦИИ-УТИЛИТЫ) ===
%>
