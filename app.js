// RACI Matrix Application
class RACIApp {
    constructor() {
        this.projects = []; // Array of all projects
        this.currentProjectIndex = 0; // Index of currently active project
        this.modals = {};
        this.selectedRows = new Set();
        this.pendingImport = null; // Temporary storage for import conflict resolution
        this.pendingExcelImport = null; // Temporary storage for Excel/CSV import
        this.init();
    }

    init() {
        // Initialize Bootstrap modals
        this.modals.column = new bootstrap.Modal(document.getElementById('columnModal'));
        this.modals.row = new bootstrap.Modal(document.getElementById('rowModal'));
        this.modals.page = new bootstrap.Modal(document.getElementById('pageModal'));
        this.modals.raci = new bootstrap.Modal(document.getElementById('raciModal'));
        this.modals.activityColumn = new bootstrap.Modal(document.getElementById('activityColumnModal'));
        this.modals.sections = new bootstrap.Modal(document.getElementById('sectionsModal'));
        this.modals.moveRow = new bootstrap.Modal(document.getElementById('moveRowModal'));
        this.modals.importConflict = new bootstrap.Modal(document.getElementById('importConflictModal'));
        this.modals.sheetSelection = new bootstrap.Modal(document.getElementById('sheetSelectionModal'));
        this.modals.csvExport = new bootstrap.Modal(document.getElementById('csvExportModal'));
        this.modals.multiPageExport = new bootstrap.Modal(document.getElementById('multiPageExportModal'));

        // Set up file input listeners
        document.getElementById('fileInput').addEventListener('change', (e) => this.handleFileImport(e));
        document.getElementById('excelFileInput').addEventListener('change', (e) => this.handleExcelFileImport(e));

        // Load projects from localStorage or create default
        this.loadFromLocalStorage();
        if (this.projects.length === 0) {
            this.createNewProject('New RACI Project');
        }
        this.render();
        this.updateProjectSwitcher();
    }

    // Get current project helper
    get project() {
        return this.projects[this.currentProjectIndex];
    }

    // Get current page index helper
    get currentPageIndex() {
        return this.project._currentPageIndex || 0;
    }

    set currentPageIndex(value) {
        this.project._currentPageIndex = value;
    }

    // Helper function to sort RACI values in the correct order (R, A, C, I)
    sortRACIValues(values) {
        const order = { 'R': 1, 'A': 2, 'C': 3, 'I': 4 };
        return values.sort((a, b) => (order[a] || 999) - (order[b] || 999));
    }

    // ========== PROJECT MANAGEMENT ==========
    createNewProject(projectName) {
        const newProject = {
            id: this.generateId(),
            projectName: projectName,
            pages: [],
            _currentPageIndex: 0
        };
        this.projects.push(newProject);
        this.currentProjectIndex = this.projects.length - 1;
        this.addPage('Page 1', true);
        document.getElementById('projectName').value = projectName;
        this.updateProjectSwitcher();
        return newProject;
    }

    newProject() {
        const projectName = prompt('Enter name for new project:', 'New RACI Project');
        if (!projectName) return;
        
        this.createNewProject(projectName);
        this.render();
        this.saveToLocalStorage();
    }

    switchProject(projectIndex) {
        if (projectIndex >= 0 && projectIndex < this.projects.length) {
            this.currentProjectIndex = projectIndex;
            this.selectedRows.clear();
            document.getElementById('projectName').value = this.project.projectName;
            this.render();
            this.updateProjectSwitcher();
        }
    }

    closeProject(projectIndex) {
        if (this.projects.length <= 1) {
            alert('Cannot close the last project. Create a new project first.');
            return;
        }

        const project = this.projects[projectIndex];
        if (!confirm(`Are you sure you want to close project "${project.projectName}"? Make sure you have exported it if needed.`)) {
            return;
        }

        this.projects.splice(projectIndex, 1);
        
        // Adjust current index if needed
        if (this.currentProjectIndex >= this.projects.length) {
            this.currentProjectIndex = this.projects.length - 1;
        }
        if (this.currentProjectIndex >= projectIndex && this.currentProjectIndex > 0) {
            this.currentProjectIndex--;
        }

        document.getElementById('projectName').value = this.project.projectName;
        this.render();
        this.saveToLocalStorage();
        this.updateProjectSwitcher();
    }

    updateProjectSwitcher() {
        const projectList = document.getElementById('projectList');
        const currentDisplay = document.getElementById('currentProjectDisplay');
        
        currentDisplay.textContent = this.project.projectName;
        
        projectList.innerHTML = '';
        
        this.projects.forEach((proj, index) => {
            const isActive = index === this.currentProjectIndex;
            const li = document.createElement('li');
            li.innerHTML = `
                <div class="dropdown-item d-flex justify-content-between align-items-center ${isActive ? 'active' : ''}" style="cursor: pointer;">
                    <span onclick="app.switchProject(${index})" style="flex: 1;">
                        <i class="fas fa-folder"></i> ${proj.projectName}
                    </span>
                    ${!isActive ? `<button class="btn btn-sm btn-link text-danger" onclick="event.stopPropagation(); app.closeProject(${index})" title="Close project">
                        <i class="fas fa-times"></i>
                    </button>` : ''}
                </div>
            `;
            projectList.appendChild(li);
        });
        
        // Add separator and new project button
        projectList.innerHTML += `
            <li><hr class="dropdown-divider"></li>
            <li><a class="dropdown-item" href="#" onclick="app.newProject()">
                <i class="fas fa-plus"></i> New Project
            </a></li>
        `;
    }

    loadFromLocalStorage() {
        const saved = localStorage.getItem('raciProjects');
        if (saved) {
            try {
                const data = JSON.parse(saved);
                this.projects = data.projects || [];
                this.currentProjectIndex = data.currentProjectIndex || 0;
                
                // Ensure backward compatibility
                this.projects.forEach(project => {
                    if (!project.id) {
                        project.id = this.generateId();
                    }
                    project.pages.forEach(page => {
                        if (!page.activityColumnName) {
                            page.activityColumnName = 'Activity/Task';
                        }
                        if (!page.sections) {
                            page.sections = [];
                        }
                        // Add parentId to sections if missing (backward compatibility)
                        page.sections.forEach(section => {
                            if (section.parentId === undefined) {
                                section.parentId = null;
                            }
                        });
                        // Normalize RACI values in all cells to ensure correct order
                        page.rows.forEach(row => {
                            Object.keys(row.cells).forEach(colId => {
                                const col = page.columns.find(c => c.id === colId);
                                if (col && col.type === 'stakeholder' && Array.isArray(row.cells[colId])) {
                                    row.cells[colId] = this.sortRACIValues([...row.cells[colId]]);
                                }
                            });
                        });
                    });
                });
                
                if (this.projects.length > 0) {
                    document.getElementById('projectName').value = this.project.projectName;
                }
            } catch (e) {
                console.error('Error loading from localStorage:', e);
                this.projects = [];
            }
        } else {
            // Check for old single-project storage format
            const oldProject = localStorage.getItem('raciProject');
            if (oldProject) {
                try {
                    const project = JSON.parse(oldProject);
                    project.id = this.generateId();
                    project._currentPageIndex = 0;
                    this.projects = [project];
                    this.currentProjectIndex = 0;
                    // Migrate to new format
                    this.saveToLocalStorage();
                    localStorage.removeItem('raciProject');
                } catch (e) {
                    console.error('Error migrating old project:', e);
                }
            }
        }
    }

    saveToLocalStorage() {
        try {
            const data = {
                projects: this.projects,
                currentProjectIndex: this.currentProjectIndex
            };
            localStorage.setItem('raciProjects', JSON.stringify(data));
        } catch (e) {
            console.error('Error saving to localStorage:', e);
        }
    }

    updateProjectName() {
        const newName = document.getElementById('projectName').value;
        if (newName && newName !== this.project.projectName) {
            this.project.projectName = newName;
            this.saveToLocalStorage();
            this.updateProjectSwitcher();
        }
    }

    // ========== PAGE MANAGEMENT ==========
    addPage(name = null, silent = false) {
        const pageName = name || `Page ${this.project.pages.length + 1}`;
        const newPage = {
            id: this.generateId(),
            name: pageName,
            activityColumnName: 'Activity/Task',
            sections: [],
            columns: [],
            rows: [],
            searchQuery: ''
        };
        
        this.project.pages.push(newPage);
        
        if (!silent) {
            this.currentPageIndex = this.project.pages.length - 1;
            this.render();
            this.saveToLocalStorage();
        }
        
        return newPage;
    }

    showAddPageModal() {
        document.getElementById('pageModalTitle').textContent = 'Add New Page';
        document.getElementById('pageName').value = '';
        document.getElementById('editPageIndex').value = '';
        this.modals.page.show();
    }

    toggleSidebar() {
        const sidebar = document.getElementById('pagesSidebar');
        const menuIcon = document.getElementById('menuToggleIcon');
        const mainContent = document.querySelector('.col-md-10, .col-md-12');
        
        if (sidebar.classList.contains('sidebar-collapsed')) {
            // Show sidebar
            sidebar.classList.remove('sidebar-collapsed');
            sidebar.classList.remove('col-md-0');
            sidebar.classList.add('col-md-2');
            mainContent.classList.remove('col-md-12');
            mainContent.classList.add('col-md-10');
            menuIcon.classList.remove('fa-chevron-right');
            menuIcon.classList.add('fa-chevron-left');
        } else {
            // Hide sidebar
            sidebar.classList.add('sidebar-collapsed');
            sidebar.classList.remove('col-md-2');
            sidebar.classList.add('col-md-0');
            mainContent.classList.remove('col-md-10');
            mainContent.classList.add('col-md-12');
            menuIcon.classList.remove('fa-chevron-left');
            menuIcon.classList.add('fa-chevron-right');
        }
    }

    savePage() {
        const pageName = document.getElementById('pageName').value.trim();
        const editIndex = document.getElementById('editPageIndex').value;
        
        if (!pageName) {
            alert('Please enter a page name');
            return;
        }
        
        if (editIndex !== '') {
            // Edit existing page
            this.project.pages[editIndex].name = pageName;
            this.render();
        } else {
            // Add new page
            this.addPage(pageName);
        }
        
        this.modals.page.hide();
        document.getElementById('pageName').value = '';
        document.getElementById('editPageIndex').value = '';
        this.saveToLocalStorage();
    }

    editPageName() {
        const page = this.getCurrentPage();
        document.getElementById('pageModalTitle').textContent = 'Rename Page';
        document.getElementById('pageName').value = page.name;
        document.getElementById('editPageIndex').value = this.currentPageIndex;
        this.modals.page.show();
    }

    switchPage(index) {
        this.currentPageIndex = index;
        this.selectedRows.clear();
        this.render();
    }

    deletePage() {
        if (this.project.pages.length === 1) {
            alert('Cannot delete the last page');
            return;
        }
        
        if (confirm('Are you sure you want to delete this page?')) {
            this.project.pages.splice(this.currentPageIndex, 1);
            this.currentPageIndex = Math.max(0, this.currentPageIndex - 1);
            this.render();
            this.saveToLocalStorage();
        }
    }

    getCurrentPage() {
        return this.project.pages[this.currentPageIndex];
    }

    // ========== BULK SELECTION MANAGEMENT ==========
    toggleSelectAll() {
        const selectAllCheckbox = document.getElementById('selectAll');
        const page = this.getCurrentPage();
        
        if (selectAllCheckbox.checked) {
            page.rows.forEach(row => this.selectedRows.add(row.id));
        } else {
            this.selectedRows.clear();
        }
        
        this.updateBulkActionBar();
        this.updateRowCheckboxes();
    }

    toggleRowSelection(rowId) {
        if (this.selectedRows.has(rowId)) {
            this.selectedRows.delete(rowId);
        } else {
            this.selectedRows.add(rowId);
        }
        
        this.updateBulkActionBar();
        this.updateSelectAllCheckbox();
    }

    updateBulkActionBar() {
        const bulkActionBar = document.getElementById('bulkActionBar');
        const selectedCount = document.getElementById('selectedCount');
        
        if (this.selectedRows.size > 0) {
            bulkActionBar.style.display = 'block';
            selectedCount.textContent = this.selectedRows.size;
            this.populateSectionDropdown('bulkSectionSelect');
        } else {
            bulkActionBar.style.display = 'none';
        }
    }

    updateSelectAllCheckbox() {
        const page = this.getCurrentPage();
        const selectAllCheckbox = document.getElementById('selectAll');
        
        if (!selectAllCheckbox) return;
        
        if (page.rows.length === 0) {
            selectAllCheckbox.checked = false;
            selectAllCheckbox.indeterminate = false;
        } else if (this.selectedRows.size === page.rows.length) {
            selectAllCheckbox.checked = true;
            selectAllCheckbox.indeterminate = false;
        } else if (this.selectedRows.size > 0) {
            selectAllCheckbox.checked = false;
            selectAllCheckbox.indeterminate = true;
        } else {
            selectAllCheckbox.checked = false;
            selectAllCheckbox.indeterminate = false;
        }
    }

    updateRowCheckboxes() {
        const page = this.getCurrentPage();
        page.rows.forEach(row => {
            const checkbox = document.getElementById(`rowCheckbox_${row.id}`);
            if (checkbox) {
                checkbox.checked = this.selectedRows.has(row.id);
            }
        });
    }

    clearSelection() {
        this.selectedRows.clear();
        this.updateBulkActionBar();
        this.updateSelectAllCheckbox();
        this.updateRowCheckboxes();
    }

    bulkAssignSection() {
        const page = this.getCurrentPage();
        const sectionId = document.getElementById('bulkSectionSelect').value || null;
        
        if (this.selectedRows.size === 0) {
            alert('No rows selected');
            return;
        }
        
        let updatedCount = 0;
        page.rows.forEach(row => {
            if (this.selectedRows.has(row.id)) {
                row.sectionId = sectionId;
                updatedCount++;
            }
        });
        
        this.clearSelection();
        this.render();
        this.saveToLocalStorage();
        
        const sectionName = sectionId ? 
            (page.sections.find(s => s.id === sectionId)?.name || 'No Section') : 
            'No Section';
        alert(`${updatedCount} row(s) assigned to "${sectionName}"`);
    }

    bulkDeleteRows() {
        const page = this.getCurrentPage();
        
        if (this.selectedRows.size === 0) {
            alert('No rows selected');
            return;
        }
        
        const count = this.selectedRows.size;
        const confirmMessage = `Are you sure you want to delete ${count} selected row(s)? This action cannot be undone.`;
        
        if (!confirm(confirmMessage)) {
            return;
        }
        
        // Remove selected rows
        page.rows = page.rows.filter(row => !this.selectedRows.has(row.id));
        
        this.clearSelection();
        this.render();
        this.saveToLocalStorage();
        
        alert(`${count} row(s) deleted successfully`);
    }

    // ========== SEARCH/FILTER ==========
    filterRows() {
        const page = this.getCurrentPage();
        const searchInput = document.getElementById('rowSearchInput');
        const clearBtn = document.getElementById('clearSearchBtn');
        const resultsText = document.getElementById('searchResultsText');
        
        page.searchQuery = searchInput.value.trim();
        
        // Show/hide clear button based on search query
        if (page.searchQuery) {
            clearBtn.style.display = 'inline-block';
        } else {
            clearBtn.style.display = 'none';
            resultsText.textContent = '';
        }
        
        // Re-render table with filtered rows
        this.renderTable();
        
        // Update results text if there's a search query
        if (page.searchQuery) {
            const allRowsCount = page.rows.length;
            const filteredCount = this.getFilteredRows().length;
            resultsText.textContent = `Showing ${filteredCount} of ${allRowsCount} rows`;
        }
    }
    
    clearSearch() {
        const page = this.getCurrentPage();
        const searchInput = document.getElementById('rowSearchInput');
        const clearBtn = document.getElementById('clearSearchBtn');
        const resultsText = document.getElementById('searchResultsText');
        
        searchInput.value = '';
        page.searchQuery = '';
        clearBtn.style.display = 'none';
        resultsText.textContent = '';
        
        this.renderTable();
    }
    
    getFilteredRows() {
        const page = this.getCurrentPage();
        if (!page.searchQuery) {
            return page.rows;
        }
        
        const query = page.searchQuery.toLowerCase();
        return page.rows.filter(row => 
            row.name.toLowerCase().includes(query)
        );
    }
    
    syncSearchUI() {
        const page = this.getCurrentPage();
        if (!page) return;
        
        const searchInput = document.getElementById('rowSearchInput');
        const clearBtn = document.getElementById('clearSearchBtn');
        const resultsText = document.getElementById('searchResultsText');
        
        // Initialize searchQuery if it doesn't exist (for legacy data)
        if (page.searchQuery === undefined) {
            page.searchQuery = '';
        }
        
        // Only sync input value if the input is not currently focused (user is not typing)
        // This prevents disrupting the cursor position while typing
        if (document.activeElement !== searchInput) {
            searchInput.value = page.searchQuery;
        }
        
        // Update clear button visibility
        if (page.searchQuery) {
            clearBtn.style.display = 'inline-block';
            const allRowsCount = page.rows.length;
            const filteredCount = this.getFilteredRows().length;
            resultsText.textContent = `Showing ${filteredCount} of ${allRowsCount} rows`;
        } else {
            clearBtn.style.display = 'none';
            resultsText.textContent = '';
        }
    }

    // ========== SECTION MANAGEMENT ==========
    manageSections() {
        this.renderSectionsList();
        this.modals.sections.show();
    }

    renderSectionsList() {
        const page = this.getCurrentPage();
        const sectionsList = document.getElementById('sectionsList');
        sectionsList.innerHTML = '';

        if (!page.sections || page.sections.length === 0) {
            sectionsList.innerHTML = '<div class="text-center text-muted p-3"><i class="fas fa-info-circle"></i> No sections created yet. Click "Add Section" to get started.</div>';
            return;
        }

        // Build hierarchy
        const renderSection = (section, level = 0) => {
            const index = page.sections.findIndex(s => s.id === section.id);
            const rowCount = page.rows.filter(r => r.sectionId === section.id).length;
            const childSections = page.sections.filter(s => s.parentId === section.id);
            const totalRows = rowCount + childSections.reduce((sum, child) => {
                const childRows = page.rows.filter(r => r.sectionId === child.id).length;
                return sum + childRows;
            }, 0);
            
            const item = document.createElement('div');
            item.className = 'list-group-item d-flex justify-content-between align-items-center';
            item.style.paddingLeft = `${0.75 + level * 2}rem`;
            
            const info = document.createElement('div');
            info.innerHTML = `
                <div class="d-flex align-items-center gap-2">
                    <i class="fas ${level > 0 ? 'fa-arrow-turn-up fa-rotate-90' : 'fa-layer-group'} text-primary"></i>
                    <div>
                        <strong>${section.name}</strong>
                        <span class="badge bg-primary ms-2">${rowCount} row${rowCount !== 1 ? 's' : ''}</span>
                        ${childSections.length > 0 ? `<span class="badge bg-secondary ms-1">${childSections.length} sub${childSections.length !== 1 ? 's' : ''}</span>` : ''}
                        ${section.description ? `<br><small class="text-muted">${section.description}</small>` : ''}
                    </div>
                </div>
            `;
            
            const actions = document.createElement('div');
            actions.className = 'btn-group btn-group-sm';
            actions.innerHTML = `
                <button class="btn btn-outline-success" onclick="app.addSubSection(${index})" title="Add Sub-section">
                    <i class="fas fa-plus"></i>
                </button>
                <button class="btn btn-outline-secondary" onclick="app.editSection(${index})" title="Edit Section">
                    <i class="fas fa-edit"></i>
                </button>
                <button class="btn btn-outline-danger" onclick="app.deleteSection(${index})" title="Delete Section">
                    <i class="fas fa-trash"></i>
                </button>
            `;
            
            item.appendChild(info);
            item.appendChild(actions);
            sectionsList.appendChild(item);
            
            // Render child sections
            childSections.forEach(child => renderSection(child, level + 1));
        };
        
        // Render top-level sections first
        const topLevelSections = page.sections.filter(s => !s.parentId);
        topLevelSections.forEach(section => renderSection(section));
    }

    addSectionInModal() {
        const page = this.getCurrentPage();
        const sectionName = prompt('Enter section name:', `Section ${(page.sections || []).length + 1}`);
        
        if (!sectionName) return;
        
        if (!page.sections) page.sections = [];
        
        page.sections.push({
            id: this.generateId(),
            name: sectionName.trim(),
            description: '',
            parentId: null
        });
        
        this.renderSectionsList();
        this.saveToLocalStorage();
    }
    
    addSubSection(parentIndex) {
        const page = this.getCurrentPage();
        const parentSection = page.sections[parentIndex];
        const sectionName = prompt(`Enter sub-section name under "${parentSection.name}":`, '');
        
        if (!sectionName) return;
        
        page.sections.push({
            id: this.generateId(),
            name: sectionName.trim(),
            description: '',
            parentId: parentSection.id
        });
        
        this.renderSectionsList();
        this.saveToLocalStorage();
    }

    editSection(index) {
        const page = this.getCurrentPage();
        const section = page.sections[index];
        
        const newName = prompt('Enter section name:', section.name);
        if (!newName) return;
        
        section.name = newName.trim();
        const newDesc = prompt('Enter section description (optional):', section.description || '');
        if (newDesc !== null) {
            section.description = newDesc.trim();
        }
        
        // Ask if user wants to change parent
        const changeParent = confirm('Would you like to change the parent section?');
        if (changeParent) {
            this.showMoveToParentDialog(section.id);
        }
        
        this.renderSectionsList();
        this.render();
        this.saveToLocalStorage();
    }
    
    showMoveToParentDialog(sectionId) {
        const page = this.getCurrentPage();
        const section = page.sections.find(s => s.id === sectionId);
        
        // Build parent options (exclude self and descendants to prevent circular reference)
        const descendants = this.getDescendantSections(sectionId);
        const excludeIds = [sectionId, ...descendants.map(s => s.id)];
        const availableSections = page.sections.filter(s => !excludeIds.includes(s.id));
        
        let message = 'Select parent section:\n\n0 - No Parent (Top Level)\n';
        availableSections.forEach((s, idx) => {
            const depth = this.getSectionDepth(s.id);
            const indent = '  '.repeat(depth);
            message += `${idx + 1} - ${indent}${s.name}\n`;
        });
        
        const choice = prompt(message, '0');
        if (choice === null) return;
        
        const choiceNum = parseInt(choice);
        if (choiceNum === 0) {
            section.parentId = null;
        } else if (choiceNum > 0 && choiceNum <= availableSections.length) {
            section.parentId = availableSections[choiceNum - 1].id;
        }
    }
    
    getDescendantSections(sectionId) {
        const page = this.getCurrentPage();
        const descendants = [];
        const findChildren = (parentId) => {
            const children = page.sections.filter(s => s.parentId === parentId);
            children.forEach(child => {
                descendants.push(child);
                findChildren(child.id);
            });
        };
        findChildren(sectionId);
        return descendants;
    }
    
    getSectionDepth(sectionId) {
        const page = this.getCurrentPage();
        let depth = 0;
        let current = page.sections.find(s => s.id === sectionId);
        while (current && current.parentId) {
            depth++;
            current = page.sections.find(s => s.id === current.parentId);
        }
        return depth;
    }

    deleteSection(index) {
        const page = this.getCurrentPage();
        const section = page.sections[index];
        
        // Check for child sections
        const childSections = page.sections.filter(s => s.parentId === section.id);
        if (childSections.length > 0) {
            if (!confirm(`This section has ${childSections.length} sub-section(s). Deleting it will move sub-sections to the parent level. Continue?`)) {
                return;
            }
            // Move child sections to parent's level
            childSections.forEach(child => child.parentId = section.parentId);
        }
        
        // Check if any rows are assigned to this section
        const rowsInSection = page.rows.filter(r => r.sectionId === section.id);
        if (rowsInSection.length > 0) {
            if (!confirm(`This section contains ${rowsInSection.length} row(s). Deleting it will move rows to "No Section". Continue?`)) {
                return;
            }
            // Move rows to no section
            rowsInSection.forEach(row => row.sectionId = null);
        }
        
        page.sections.splice(index, 1);
        this.renderSectionsList();
        this.render();
        this.saveToLocalStorage();
    }

    populateSectionDropdown(selectId) {
        const page = this.getCurrentPage();
        const select = document.getElementById(selectId);
        select.innerHTML = '<option value="">No Section</option>';
        
        if (!page.sections || page.sections.length === 0) return;
        
        // Render sections hierarchically
        const renderSectionOption = (section, level = 0) => {
            const indent = '\u00A0\u00A0'.repeat(level * 2); // Non-breaking spaces for indentation
            const prefix = level > 0 ? '└─ ' : '';
            const option = document.createElement('option');
            option.value = section.id;
            option.textContent = indent + prefix + section.name;
            select.appendChild(option);
            
            // Add child sections
            const children = page.sections.filter(s => s.parentId === section.id);
            children.forEach(child => renderSectionOption(child, level + 1));
        };
        
        // Render top-level sections
        const topLevel = page.sections.filter(s => !s.parentId);
        topLevel.forEach(section => renderSectionOption(section));
    }
    
    populateSectionDropdownOld(selectId) {
        const page = this.getCurrentPage();
        const select = document.getElementById(selectId);
        select.innerHTML = '<option value="">No Section</option>';
        
        if (page.sections && page.sections.length > 0) {
            page.sections.forEach(section => {
                const option = document.createElement('option');
                option.value = section.id;
                option.textContent = section.name;
                select.appendChild(option);
            });
        }
    }

    editActivityColumnName() {
        const page = this.getCurrentPage();
        document.getElementById('activityColumnName').value = page.activityColumnName || 'Activity/Task';
        this.modals.activityColumn.show();
    }

    saveActivityColumnName() {
        const page = this.getCurrentPage();
        const name = document.getElementById('activityColumnName').value.trim();
        
        if (!name) {
            alert('Please enter a column name');
            return;
        }
        
        page.activityColumnName = name;
        this.modals.activityColumn.hide();
        this.render();
        this.saveToLocalStorage();
    }

    // ========== COLUMN MANAGEMENT ==========
    addColumn() {
        document.getElementById('columnName').value = '';
        document.getElementById('columnType').value = 'stakeholder';
        document.getElementById('columnDescription').value = '';
        document.getElementById('columnId').value = '';
        document.getElementById('columnIndex').value = '';
        document.getElementById('deleteColumnBtn').style.display = 'none';
        this.modals.column.show();
    }

    editColumn(index) {
        const page = this.getCurrentPage();
        const column = page.columns[index];
        
        document.getElementById('columnName').value = column.name;
        document.getElementById('columnType').value = column.type;
        document.getElementById('columnDescription').value = column.description || '';
        document.getElementById('columnId').value = column.id;
        document.getElementById('columnIndex').value = index;
        document.getElementById('deleteColumnBtn').style.display = 'block';
        this.modals.column.show();
    }

    saveColumn() {
        const page = this.getCurrentPage();
        const name = document.getElementById('columnName').value.trim();
        const type = document.getElementById('columnType').value;
        const description = document.getElementById('columnDescription').value.trim();
        const columnId = document.getElementById('columnId').value;
        const columnIndex = document.getElementById('columnIndex').value;

        if (!name) {
            alert('Please enter a column name');
            return;
        }

        if (columnId) {
            // Edit existing column
            const column = page.columns[columnIndex];
            column.name = name;
            column.type = type;
            column.description = description;
        } else {
            // Add new column
            const newColumn = {
                id: this.generateId(),
                name: name,
                type: type,
                description: description
            };
            page.columns.push(newColumn);
        }

        this.modals.column.hide();
        this.render();
        this.saveToLocalStorage();
    }

    deleteColumn() {
        const columnIndex = parseInt(document.getElementById('columnIndex').value);
        const page = this.getCurrentPage();
        const column = page.columns[columnIndex];

        if (confirm(`Delete column "${column.name}"?`)) {
            // Remove column from all rows
            page.rows.forEach(row => {
                delete row.cells[column.id];
            });
            
            page.columns.splice(columnIndex, 1);
            this.modals.column.hide();
            this.render();
            this.saveToLocalStorage();
        }
    }

    // ========== ROW MANAGEMENT ==========
    addRow() {
        document.getElementById('rowName').value = '';
        document.getElementById('rowDescription').value = '';
        document.getElementById('bulkRowData').value = '';
        document.getElementById('rowId').value = '';
        document.getElementById('rowIndex').value = '';
        document.getElementById('bulkRowMode').checked = false;
        document.getElementById('deleteRowBtn').style.display = 'none';
        document.getElementById('singleRowInputs').style.display = 'block';
        document.getElementById('bulkRowInputs').style.display = 'none';
        this.populateSectionDropdown('rowSection');
        this.populateSectionDropdown('bulkRowSection');
        document.getElementById('rowSection').value = '';
        this.modals.row.show();
    }

    toggleBulkRowMode() {
        const bulkMode = document.getElementById('bulkRowMode').checked;
        const singleInputs = document.getElementById('singleRowInputs');
        const bulkInputs = document.getElementById('bulkRowInputs');
        
        if (bulkMode) {
            singleInputs.style.display = 'none';
            bulkInputs.style.display = 'block';
            // Populate section dropdown for bulk mode
            this.populateSectionDropdown('bulkRowSection');
        } else {
            singleInputs.style.display = 'block';
            bulkInputs.style.display = 'none';
        }
    }

    editRow(index) {
        const page = this.getCurrentPage();
        const row = page.rows[index];
        
        document.getElementById('rowName').value = row.name;
        document.getElementById('rowDescription').value = row.description || '';
        document.getElementById('rowId').value = row.id;
        document.getElementById('rowIndex').value = index;
        document.getElementById('deleteRowBtn').style.display = 'block';
        document.getElementById('bulkRowMode').checked = false;
        document.getElementById('singleRowInputs').style.display = 'block';
        document.getElementById('bulkRowInputs').style.display = 'none';
        this.populateSectionDropdown('rowSection');
        document.getElementById('rowSection').value = row.sectionId || '';
        this.modals.row.show();
    }

    saveRow() {
        const page = this.getCurrentPage();
        const bulkMode = document.getElementById('bulkRowMode').checked;
        const rowId = document.getElementById('rowId').value;
        const rowIndex = document.getElementById('rowIndex').value;

        if (bulkMode) {
            // Bulk add mode
            const bulkData = document.getElementById('bulkRowData').value.trim();
            const sectionId = document.getElementById('bulkRowSection').value;
            
            if (!bulkData) {
                alert('Please paste some tasks (one per line)');
                return;
            }

            // Split by lines and filter empty lines
            const lines = bulkData.split('\n')
                .map(line => line.trim())
                .filter(line => line.length > 0);

            if (lines.length === 0) {
                alert('No valid tasks found. Please enter one task per line.');
                return;
            }

            // Create a row for each line
            let addedCount = 0;
            lines.forEach(line => {
                const newRow = {
                    id: this.generateId(),
                    name: line,
                    description: '',
                    sectionId: sectionId || null,
                    cells: {}
                };
                page.rows.push(newRow);
                addedCount++;
            });

            this.modals.row.hide();
            this.render();
            this.saveToLocalStorage();
            
            // Show section name in confirmation if assigned
            let message = `Successfully added ${addedCount} row${addedCount > 1 ? 's' : ''}!`;
            if (sectionId) {
                const section = page.sections.find(s => s.id === sectionId);
                if (section) {
                    message += ` All rows assigned to "${section.name}".`;
                }
            }
            alert(message);
        } else {
            // Single row mode
            const name = document.getElementById('rowName').value.trim();
            const description = document.getElementById('rowDescription').value.trim();
            const sectionId = document.getElementById('rowSection').value;

            if (!name) {
                alert('Please enter an activity/task name');
                return;
            }

            if (rowId) {
                // Edit existing row
                const row = page.rows[rowIndex];
                row.name = name;
                row.description = description;
                row.sectionId = sectionId || null;
            } else {
                // Add new row
                const newRow = {
                    id: this.generateId(),
                    name: name,
                    description: description,
                    sectionId: sectionId || null,
                    cells: {}
                };
                page.rows.push(newRow);
            }

            this.modals.row.hide();
            this.render();
            this.saveToLocalStorage();
        }
    }

    deleteRow() {
        const rowIndex = parseInt(document.getElementById('rowIndex').value);
        const page = this.getCurrentPage();
        const row = page.rows[rowIndex];

        if (confirm(`Delete row "${row.name}"?`)) {
            page.rows.splice(rowIndex, 1);
            this.modals.row.hide();
            this.render();
            this.saveToLocalStorage();
        }
    }

    moveRowToSection(rowId) {
        document.getElementById('moveRowId').value = rowId;
        this.populateSectionDropdown('moveToSection');
        
        const page = this.getCurrentPage();
        const row = page.rows.find(r => r.id === rowId);
        document.getElementById('moveToSection').value = row.sectionId || '';
        
        this.modals.moveRow.show();
    }

    saveRowMove() {
        const page = this.getCurrentPage();
        const rowId = document.getElementById('moveRowId').value;
        const sectionId = document.getElementById('moveToSection').value;
        
        const row = page.rows.find(r => r.id === rowId);
        if (row) {
            row.sectionId = sectionId || null;
            this.modals.moveRow.hide();
            this.render();
            this.saveToLocalStorage();
        }
    }
    
    convertRowToSection(rowId) {
        const page = this.getCurrentPage();
        const row = page.rows.find(r => r.id === rowId);
        
        if (!row) return;
        
        // Build warning message
        const hasData = Object.keys(row.cells || {}).some(colId => {
            const cellValue = row.cells[colId];
            return cellValue && cellValue.length > 0;
        });
        
        let warningMsg = `Convert row "${row.name}" into a section?\n\n`;
        warningMsg += `This will:\n`;
        warningMsg += `• Create a new section header with this name\n`;
        warningMsg += `• Remove this row from the matrix\n`;
        if (hasData) {
            warningMsg += `• ⚠️ DELETE all RACI assignments for this row\n`;
        }
        if (row.sectionId) {
            const parentSection = page.sections?.find(s => s.id === row.sectionId);
            if (parentSection) {
                warningMsg += `• Maintain hierarchy under section "${parentSection.name}"\n`;
            }
        }
        warningMsg += `\nThis action cannot be undone. Continue?`;
        
        if (!confirm(warningMsg)) return;
        
        // Initialize sections array if needed
        if (!page.sections) page.sections = [];
        
        // Create new section from row
        const newSection = {
            id: this.generateId(),
            name: row.name,
            description: row.description || '',
            parentId: row.sectionId || null  // Preserve hierarchy
        };
        
        page.sections.push(newSection);
        
        // Remove the row
        const rowIndex = page.rows.findIndex(r => r.id === rowId);
        if (rowIndex !== -1) {
            page.rows.splice(rowIndex, 1);
        }
        
        // Update current page index if needed
        if (page._currentPageIndex >= page.rows.length && page.rows.length > 0) {
            page._currentPageIndex = page.rows.length - 1;
        }
        
        this.render();
        this.saveToLocalStorage();
        
        // Show success message
        alert(`✓ Row "${newSection.name}" has been converted to a section.\n\nYou can now assign rows to this section or create sub-sections under it.`);
    }

    // ========== RACI CELL MANAGEMENT ==========
    // Toggle individual RACI value by clicking on badge
    toggleRACIValue(rowId, colId, letter) {
        const page = this.getCurrentPage();
        const row = page.rows.find(r => r.id === rowId);
        
        if (!row.cells[colId]) {
            row.cells[colId] = [];
        }
        
        const currentValues = row.cells[colId];
        const index = currentValues.indexOf(letter);
        
        if (index > -1) {
            // Remove the letter (deselect)
            currentValues.splice(index, 1);
        } else {
            // Add the letter (select)
            currentValues.push(letter);
        }
        
        // Sort values in RACI order (R, A, C, I)
        this.sortRACIValues(currentValues);
        
        this.saveToLocalStorage();
        this.renderTable();
    }

    editRACICell(rowId, colId) {
        const page = this.getCurrentPage();
        const column = page.columns.find(c => c.id === colId);
        
        // Only allow RACI editing for stakeholder columns
        if (column.type !== 'stakeholder') {
            return;
        }

        const row = page.rows.find(r => r.id === rowId);
        const currentValues = row.cells[colId] || [];

        // Set checkboxes
        document.getElementById('raciR').checked = currentValues.includes('R');
        document.getElementById('raciA').checked = currentValues.includes('A');
        document.getElementById('raciC').checked = currentValues.includes('C');
        document.getElementById('raciI').checked = currentValues.includes('I');

        document.getElementById('currentRowId').value = rowId;
        document.getElementById('currentColId').value = colId;

        this.modals.raci.show();
    }

    saveRACICell() {
        const page = this.getCurrentPage();
        const rowId = document.getElementById('currentRowId').value;
        const colId = document.getElementById('currentColId').value;

        const row = page.rows.find(r => r.id === rowId);
        const values = [];

        if (document.getElementById('raciR').checked) values.push('R');
        if (document.getElementById('raciA').checked) values.push('A');
        if (document.getElementById('raciC').checked) values.push('C');
        if (document.getElementById('raciI').checked) values.push('I');

        if (values.length > 0) {
            // Values are already in RACI order due to the if statement order above,
            // but sort anyway for consistency
            row.cells[colId] = this.sortRACIValues(values);
        } else {
            delete row.cells[colId];
        }

        this.modals.raci.hide();
        this.render();
        this.saveToLocalStorage();
    }

    editInfoCell(rowId, colId) {
        const page = this.getCurrentPage();
        const row = page.rows.find(r => r.id === rowId);
        const column = page.columns.find(c => c.id === colId);

        const currentValue = row.cells[colId] ? row.cells[colId][0] : '';
        const newValue = prompt(`Enter value for "${column.name}":`, currentValue);

        if (newValue !== null) {
            if (newValue.trim()) {
                row.cells[colId] = [newValue.trim()];
            } else {
                delete row.cells[colId];
            }
            this.render();
            this.saveToLocalStorage();
        }
    }

    // ========== RENDERING ==========
    render() {
        this.updateProjectName();
        this.renderPagesList();
        this.renderTable();
    }

    renderPagesList() {
        const pagesList = document.getElementById('pagesList');
        pagesList.innerHTML = '';

        this.project.pages.forEach((page, index) => {
            const item = document.createElement('a');
            item.href = '#';
            item.className = `list-group-item list-group-item-action ${index === this.currentPageIndex ? 'active' : ''}`;
            
            const nameSpan = document.createElement('span');
            nameSpan.textContent = page.name;
            nameSpan.style.flexGrow = '1';
            nameSpan.onclick = (e) => {
                e.preventDefault();
                this.switchPage(index);
            };
            
            const editBtn = document.createElement('button');
            editBtn.className = 'btn btn-sm btn-link text-white p-0';
            editBtn.innerHTML = '<i class="fas fa-edit"></i>';
            editBtn.onclick = (e) => {
                e.preventDefault();
                e.stopPropagation();
                this.currentPageIndex = index;
                this.editPageName();
            };
            
            item.style.display = 'flex';
            item.style.justifyContent = 'space-between';
            item.style.alignItems = 'center';
            item.appendChild(nameSpan);
            item.appendChild(editBtn);
            
            pagesList.appendChild(item);
        });
    }

    renderTable() {
        const page = this.getCurrentPage();
        if (!page) return;

        document.getElementById('currentPageName').textContent = page.name;

        // Render header
        const header = document.getElementById('tableHeader');
        const activityColName = page.activityColumnName || 'Activity/Task';
        header.innerHTML = `
            <th class="checkbox-column">
                <input type="checkbox" id="selectAll" onclick="app.toggleSelectAll()" title="Select All">
            </th>
            <th class="activity-column">
                <span>${activityColName}</span>
                <button class="btn btn-sm btn-link p-0 ms-2" onclick="app.editActivityColumnName()" title="Rename this column">
                    <i class="fas fa-edit text-secondary"></i>
                </button>
            </th>
        `;

        page.columns.forEach((col, index) => {
            const th = document.createElement('th');
            th.className = `column-${col.type}`;
            th.innerHTML = `
                <div class="d-flex justify-content-between align-items-center">
                    <span title="${col.description || ''}">${col.name}</span>
                    <button class="btn btn-sm btn-link" onclick="app.editColumn(${index})">
                        <i class="fas fa-edit"></i>
                    </button>
                </div>
                <div class="column-type-badge">
                    <span class="badge ${col.type === 'stakeholder' ? 'bg-primary' : 'bg-info'}">${col.type}</span>
                </div>
            `;
            header.appendChild(th);
        });

        // Render body
        const tbody = document.getElementById('tableBody');
        tbody.innerHTML = '';

        // Get filtered rows based on search query
        const filteredRows = this.getFilteredRows();

        if (page.rows.length === 0) {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td colspan="${page.columns.length + 2}" class="text-center text-muted">
                    <i class="fas fa-info-circle"></i> No activities/tasks yet. Click "Add Row" to get started.
                </td>
            `;
            tbody.appendChild(tr);
            this.updateBulkActionBar();
            this.updateSelectAllCheckbox();
            return;
        }

        if (filteredRows.length === 0) {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td colspan="${page.columns.length + 2}" class="text-center text-muted">
                    <i class="fas fa-search"></i> No rows match your search. Try a different search term.
                </td>
            `;
            tbody.appendChild(tr);
            this.updateBulkActionBar();
            this.updateSelectAllCheckbox();
            return;
        }

        // Group rows by section
        const sections = page.sections || [];
        const rowsBySection = new Map();
        
        // Initialize with sections
        sections.forEach(section => {
            rowsBySection.set(section.id, []);
        });
        
        // Add "No Section" group
        rowsBySection.set(null, []);
        
        // Distribute rows into sections
        filteredRows.forEach((row, rowIndex) => {
            const sectionId = row.sectionId || null;
            if (!rowsBySection.has(sectionId)) {
                rowsBySection.set(sectionId, []);
            }
            rowsBySection.get(sectionId).push({ row, rowIndex });
        });

        // Render sections hierarchically
        const renderSectionAndRows = (sectionId, level = 0) => {
            const items = rowsBySection.get(sectionId) || [];
            
            // Render section header
            if (sectionId !== null && items.length > 0) {
                const section = sections.find(s => s.id === sectionId);
                const sectionRow = document.createElement('tr');
                sectionRow.className = 'section-header-row';
                const indent = level * 1.5; // rem units for indentation
                sectionRow.innerHTML = `
                    <td class="section-header-checkbox"></td>
                    <td colspan="${page.columns.length + 1}" class="section-header" style="padding-left: ${0.75 + indent}rem;">
                        <i class="fas ${level > 0 ? 'fa-arrow-turn-up fa-rotate-90' : 'fa-layer-group'}"></i> ${section.name}
                        ${section.description ? `<span class="text-muted ms-2">— ${section.description}</span>` : ''}
                    </td>
                `;
                tbody.appendChild(sectionRow);
            }
            
            // Render rows in this section
            items.forEach(({ row, rowIndex }) => {
                const tr = document.createElement('tr');
                
                // Checkbox cell
                const checkboxTd = document.createElement('td');
                checkboxTd.className = 'checkbox-column';
                checkboxTd.innerHTML = `
                    <input type="checkbox" 
                           id="rowCheckbox_${row.id}" 
                           ${this.selectedRows.has(row.id) ? 'checked' : ''}
                           onclick="app.toggleRowSelection('${row.id}')">
                `;
                tr.appendChild(checkboxTd);
                
                // Activity/Task name cell
                const nameTd = document.createElement('td');
                nameTd.className = 'activity-column';
                
                // Get section name if row has one
                let sectionBadge = '';
                if (row.sectionId) {
                    const section = sections.find(s => s.id === row.sectionId);
                    if (section) {
                        sectionBadge = `<span class="badge bg-secondary ms-2" style="font-size: 0.7rem;">${section.name}</span>`;
                    }
                }
                
                // Add indentation if row belongs to a section
                const rowIndent = sectionId ? (level + 1) * 1.5 : 0; // Extra level for rows under sections
                if (sectionId) {
                    nameTd.style.paddingLeft = `${0.75 + rowIndent}rem`;
                }
                
                nameTd.innerHTML = `
                    <div class="d-flex justify-content-between align-items-center">
                        <span title="${row.description || ''}">
                            ${row.name}
                            ${sectionBadge}
                        </span>
                        <div class="btn-group btn-group-sm row-actions">
                            <button class="btn btn-outline-secondary btn-sm" onclick="app.moveRowToSection('${row.id}')" title="Move to Section">
                                <i class="fas fa-arrows-alt"></i>
                            </button>
                            <button class="btn btn-outline-warning btn-sm" onclick="app.convertRowToSection('${row.id}')" title="Convert to Section">
                                <i class="fas fa-layer-group"></i>
                            </button>
                            <button class="btn btn-outline-primary btn-sm" onclick="app.editRow(${rowIndex})" title="Edit Row">
                                <i class="fas fa-edit"></i>
                            </button>
                        </div>
                    </div>
                `;
                tr.appendChild(nameTd);

            // RACI cells
            page.columns.forEach(col => {
                const td = document.createElement('td');
                td.className = 'raci-cell';
                
                const cellValue = row.cells[col.id] || [];
                
                if (col.type === 'stakeholder') {
                    // Always display R A C I with colored badges - each badge is individually clickable
                    const raciLetters = ['R', 'A', 'C', 'I'];
                    const badgesHtml = raciLetters.map(letter => {
                        const isSelected = cellValue.includes(letter);
                        const className = isSelected ? `raci-badge raci-${letter.toLowerCase()}-selected` : 'raci-badge raci-unselected';
                        return `<span class="${className}" onclick="app.toggleRACIValue('${row.id}', '${col.id}', '${letter}')" style="cursor: pointer;" title="Click to toggle ${letter}">${letter}</span>`;
                    }).join(' ');
                    
                    td.innerHTML = badgesHtml;
                } else {
                    td.innerHTML = cellValue[0] || '';
                    td.onclick = () => this.editInfoCell(row.id, col.id);
                    td.style.cursor = 'pointer';
                    td.className += ' info-cell';
                }
                
                tr.appendChild(td);
            });

                tbody.appendChild(tr);
            });
            
            // Render child sections recursively
            if (sectionId !== null) {
                const childSections = sections.filter(s => s.parentId === sectionId);
                childSections.forEach(child => renderSectionAndRows(child.id, level + 1));
            }
        };
        
        // Start with "No Section" rows, then top-level sections
        renderSectionAndRows(null, 0);
        const topLevelSections = sections.filter(s => !s.parentId);
        topLevelSections.forEach(section => renderSectionAndRows(section.id, 0));
        
        // Update bulk action bar and checkboxes after rendering
        this.updateBulkActionBar();
        this.updateSelectAllCheckbox();
        this.syncSearchUI();
    }

    // ========== EXPORT/IMPORT ==========
    
    // Encryption key (in production, this could be user-configurable)
    getEncryptionKey() {
        // This is a fixed key for the app. In a real scenario, you might want to make this configurable
        const keyMaterial = 'RACIMatrix-Secure-Key-2026';
        return keyMaterial;
    }
    
    async encryptData(jsonString) {
        try {
            const encoder = new TextEncoder();
            const data = encoder.encode(jsonString);
            const keyMaterial = this.getEncryptionKey();
            
            // Generate a key from the password
            const keyBytes = encoder.encode(keyMaterial);
            const keyHash = await crypto.subtle.digest('SHA-256', keyBytes);
            const key = await crypto.subtle.importKey(
                'raw',
                keyHash,
                { name: 'AES-GCM' },
                false,
                ['encrypt']
            );
            
            // Generate a random IV
            const iv = crypto.getRandomValues(new Uint8Array(12));
            
            // Encrypt the data
            const encryptedData = await crypto.subtle.encrypt(
                { name: 'AES-GCM', iv: iv },
                key,
                data
            );
            
            // Combine IV and encrypted data
            const combined = new Uint8Array(iv.length + encryptedData.byteLength);
            combined.set(iv, 0);
            combined.set(new Uint8Array(encryptedData), iv.length);
            
            // Convert to base64
            return btoa(String.fromCharCode(...combined));
        } catch (err) {
            console.error('Encryption error:', err);
            throw new Error('Failed to encrypt data');
        }
    }
    
    async decryptData(encryptedBase64) {
        try {
            const encoder = new TextEncoder();
            const keyMaterial = this.getEncryptionKey();
            
            // Generate the same key
            const keyBytes = encoder.encode(keyMaterial);
            const keyHash = await crypto.subtle.digest('SHA-256', keyBytes);
            const key = await crypto.subtle.importKey(
                'raw',
                keyHash,
                { name: 'AES-GCM' },
                false,
                ['decrypt']
            );
            
            // Decode base64
            const combined = new Uint8Array(
                atob(encryptedBase64).split('').map(c => c.charCodeAt(0))
            );
            
            // Extract IV and encrypted data
            const iv = combined.slice(0, 12);
            const encryptedData = combined.slice(12);
            
            // Decrypt
            const decryptedData = await crypto.subtle.decrypt(
                { name: 'AES-GCM', iv: iv },
                key,
                encryptedData
            );
            
            // Convert back to string
            const decoder = new TextDecoder();
            return decoder.decode(decryptedData);
        } catch (err) {
            console.error('Decryption error:', err);
            throw new Error('Failed to decrypt data. File may be corrupted or invalid.');
        }
    }
    
    async exportJSON() {
        const dataStr = JSON.stringify(this.project, null, 2);
        
        try {
            // Encrypt the JSON data
            const encryptedData = await this.encryptData(dataStr);
            
            // Create blob with encrypted data
            const dataBlob = new Blob([encryptedData], { type: 'application/octet-stream' });
            const url = URL.createObjectURL(dataBlob);
            const link = document.createElement('a');
            link.href = url;
            link.download = `${this.project.projectName || 'raci-matrix'}.raci`;
            link.click();
            URL.revokeObjectURL(url);
        } catch (err) {
            alert('Error exporting project: ' + err.message);
        }
    }
    
    async saveAsRaci() {
        // Prompt for new project name
        const newName = prompt('Enter new project name:', this.project.projectName || 'raci-matrix');
        if (!newName) return; // User cancelled
        
        const originalName = this.project.projectName;
        this.project.projectName = newName;
        
        try {
            await this.exportJSON();
        } catch (err) {
            // Restore original name if save failed
            this.project.projectName = originalName;
            alert('Error saving project: ' + err.message);
        }
    }
    
    closeProject() {
        if (confirm('Close current project? Any unsaved changes will be lost.')) {
            this.newProject();
        }
    }
    
    exportToJSON() {
        // Plain JSON export (unencrypted)
        const dataStr = JSON.stringify(this.project, null, 2);
        const dataBlob = new Blob([dataStr], { type: 'application/json' });
        const url = URL.createObjectURL(dataBlob);
        const link = document.createElement('a');
        link.href = url;
        link.download = `${this.project.projectName || 'raci-matrix'}.json`;
        link.click();
        URL.revokeObjectURL(url);
    }
    
    exportToCSV() {
        const pages = this.project.pages;
        if (pages.length === 0) {
            alert('No pages to export.');
            return;
        }
        
        // Populate the page dropdown
        const select = document.getElementById('csvPageSelect');
        select.innerHTML = '';
        pages.forEach((page, index) => {
            const option = document.createElement('option');
            option.value = index;
            option.textContent = page.name;
            select.appendChild(option);
        });
        
        // Show the modal
        this.modals.csvExport.show();
    }
    
    confirmCSVExport() {
        const select = document.getElementById('csvPageSelect');
        const pageIndex = parseInt(select.value);
        const page = this.project.pages[pageIndex];
        
        if (page) {
            this.exportPageToCSV(page);
            this.modals.csvExport.hide();
        } else {
            alert('Please select a valid page.');
        }
    }
    
    exportPageToCSV(page) {
        // Build CSV content
        const rows = [];
        
        // Header row: Activity column + all other columns
        const headers = [page.activityColumnName || 'Activity'];
        page.columns.forEach(col => headers.push(col.name));
        rows.push(headers);
        
        // Get all non-section rows
        const dataRows = page.rows.filter(row => !row.isSection);
        
        // Data rows
        dataRows.forEach(row => {
            const csvRow = [row.name];
            page.columns.forEach(col => {
                const cellValue = row.cells[col.id];
                if (col.type === 'stakeholder') {
                    // Join RACI values with comma
                    csvRow.push(cellValue ? cellValue.join(',') : '');
                } else {
                    // Information column
                    csvRow.push(cellValue ? cellValue.join(', ') : '');
                }
            });
            rows.push(csvRow);
        });
        
        // Convert to CSV string
        const csvContent = rows.map(row => 
            row.map(cell => {
                // Escape quotes and wrap in quotes if contains comma, quote, or newline
                const cellStr = String(cell);
                if (cellStr.includes(',') || cellStr.includes('"') || cellStr.includes('\n')) {
                    return '"' + cellStr.replace(/"/g, '""') + '"';
                }
                return cellStr;
            }).join(',')
        ).join('\n');
        
        // Download
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = `${page.name || 'page'}.csv`;
        link.click();
        URL.revokeObjectURL(url);
    }
    
    showMultiPageExportModal(exportType) {
        // Store the export type for later use
        this.currentExportType = exportType;
        
        // Update modal title and button based on export type
        const titles = {
            'excel': '<i class="fas fa-file-excel"></i> Export to Excel',
            'pdf': '<i class="fas fa-file-pdf"></i> Export to PDF',
            'word': '<i class="fas fa-file-word"></i> Export to Word',
            'powerpoint': '<i class="fas fa-file-powerpoint"></i> Export to PowerPoint'
        };
        
        document.getElementById('multiPageExportTitle').innerHTML = titles[exportType] || 'Export';
        
        // Populate pages list
        const pagesList = document.getElementById('exportPagesList');
        pagesList.innerHTML = '';
        
        if (this.project.pages.length === 0) {
            pagesList.innerHTML = '<div class="list-group-item text-muted">No pages available to export.</div>';
            return;
        }
        
        this.project.pages.forEach((page, index) => {
            const item = document.createElement('div');
            item.className = 'list-group-item';
            item.innerHTML = `
                <div class="form-check">
                    <input class="form-check-input export-page-checkbox" type="checkbox" value="${index}" 
                           id="exportPage_${index}" checked>
                    <label class="form-check-label" for="exportPage_${index}">
                        ${page.name}
                    </label>
                </div>
            `;
            pagesList.appendChild(item);
        });
        
        // Show the modal
        this.modals.multiPageExport.show();
    }
    
    toggleAllExportPages(checked) {
        document.querySelectorAll('.export-page-checkbox').forEach(checkbox => {
            checkbox.checked = checked;
        });
    }
    
    confirmMultiPageExport() {
        // Get selected page indices
        const selectedIndices = [];
        document.querySelectorAll('.export-page-checkbox:checked').forEach(checkbox => {
            selectedIndices.push(parseInt(checkbox.value));
        });
        
        if (selectedIndices.length === 0) {
            alert('Please select at least one page to export.');
            return;
        }
        
        // Get indentation setting
        const useIndent = document.getElementById('exportWithIndent').checked;
        
        // Hide modal
        this.modals.multiPageExport.hide();
        
        // Call appropriate export method with selected pages and indent setting
        switch(this.currentExportType) {
            case 'excel':
                this.performExcelExport(selectedIndices, useIndent);
                break;
            case 'pdf':
                this.performPDFExport(selectedIndices, useIndent);
                break;
            case 'word':
                this.performWordExport(selectedIndices, useIndent);
                break;
            case 'powerpoint':
                this.performPowerPointExport(selectedIndices, useIndent);
                break;
        }
    }

    importProject() {
        document.getElementById('fileInput').click();
    }
    
    importFromExcel() {
        document.getElementById('excelFileInput').click();
    }
    
    async handleExcelFileImport(event) {
        const file = event.target.files[0];
        if (!file) return;
        
        try {
            const data = await file.arrayBuffer();
            const workbook = XLSX.read(data, { type: 'array' });
            
            // Store workbook data for later processing
            this.pendingExcelImport = {
                workbook: workbook,
                fileName: file.name
            };
            
            // Display file name
            document.getElementById('excelFileName').textContent = file.name;
            
            // Reset import mode to default (new project)
            document.getElementById('importModeNewProject').checked = true;
            
            // Populate sheet list with checkboxes and per-sheet configuration
            const sheetsList = document.getElementById('sheetsList');
            sheetsList.innerHTML = '';
            
            workbook.SheetNames.forEach((sheetName, index) => {
                const sheet = workbook.Sheets[sheetName];
                const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: true });
                const rowCount = jsonData.filter(row => row && row.length > 0).length;
                const colCount = jsonData[0] ? jsonData[0].length : 0;
                
                // Detect RACI columns for preview (using default row 1 headers)
                const columnTypes = this.detectRACIColumns(jsonData);
                const stakeholderCount = columnTypes.filter(c => c.type === 'stakeholder').length;
                const infoCount = columnTypes.filter(c => c.type === 'information').length - 1; // -1 for activity column
                
                const item = document.createElement('div');
                item.className = 'list-group-item p-0';
                item.innerHTML = `
                    <div class="p-3">
                        <div class="d-flex justify-content-between align-items-start">
                            <div class="form-check flex-grow-1">
                                <input class="form-check-input" type="checkbox" value="${index}" id="sheet_${index}" ${index === 0 ? 'checked' : ''}>
                                <label class="form-check-label w-100" for="sheet_${index}">
                                    <div class="d-flex justify-content-between align-items-center">
                                        <div>
                                            <strong>${sheetName}</strong>
                                            <span class="badge bg-secondary ms-2">${rowCount} rows</span>
                                            <span class="badge bg-info ms-1">${colCount} cols</span>
                                        </div>
                                        <div>
                                            <span class="badge bg-primary" title="Stakeholder columns detected">${stakeholderCount} <i class="fas fa-users"></i></span>
                                            <span class="badge bg-success ms-1" title="Information columns detected">${infoCount} <i class="fas fa-info-circle"></i></span>
                                        </div>
                                    </div>
                                </label>
                            </div>
                            <button class="btn btn-sm btn-outline-secondary ms-2" type="button" 
                                    data-bs-toggle="collapse" data-bs-target="#config_${index}" 
                                    aria-expanded="false" aria-controls="config_${index}"
                                    title="Configure sheet settings">
                                <i class="fas fa-cog"></i>
                            </button>
                        </div>
                        <div class="collapse mt-3" id="config_${index}">
                            <div class="card card-body bg-light">
                                <h6 class="mb-3"><i class="fas fa-sliders-h"></i> Sheet Configuration</h6>
                                <div class="mb-2">
                                    <div class="form-check">
                                        <input class="form-check-input" type="checkbox" id="hasHeaders_${index}" 
                                               checked onchange="app.toggleSheetHeaders(${index})">
                                        <label class="form-check-label" for="hasHeaders_${index}">
                                            <i class="fas fa-check-circle"></i> Has Headers
                                        </label>
                                    </div>
                                </div>
                                <div class="row g-2">
                                    <div class="col-md-6">
                                        <label for="headerRow_${index}" class="form-label text-sm">Header Row:</label>
                                        <input type="number" class="form-control form-control-sm" 
                                               id="headerRow_${index}" value="1" min="1" max="100"
                                               onchange="app.updateSheetDataRow(${index})">
                                        <small class="text-muted">Row with column headers</small>
                                    </div>
                                    <div class="col-md-6">
                                        <label for="dataRow_${index}" class="form-label text-sm">Data Starts:</label>
                                        <input type="number" class="form-control form-control-sm" 
                                               id="dataRow_${index}" value="2" min="2" max="100">
                                        <small class="text-muted">First row of data</small>
                                    </div>
                                </div>
                                <div class="mt-2">
                                    <small class="text-muted">
                                        <i class="fas fa-info-circle"></i> 
                                        If this sheet has metadata/title rows at the top, set header row accordingly.
                                        <strong>Example:</strong> Headers in row 4, data in row 5.
                                    </small>
                                </div>
                            </div>
                        </div>
                    </div>
                `;
                sheetsList.appendChild(item);
            });
            
            // Show the sheet selection modal
            this.modals.sheetSelection.show();
            
        } catch (error) {
            alert('Error reading Excel/CSV file: ' + error.message);
        }
        
        // Reset file input
        event.target.value = '';
    }
    
    updateSheetDataRow(sheetIndex) {
        const headerRow = parseInt(document.getElementById(`headerRow_${sheetIndex}`).value) || 1;
        const dataRowField = document.getElementById(`dataRow_${sheetIndex}`);
        
        // Auto-update data start row to be header row + 1
        dataRowField.value = headerRow + 1;
        dataRowField.min = headerRow + 1;
    }
    
    toggleSheetHeaders(sheetIndex) {
        const hasHeadersCheckbox = document.getElementById(`hasHeaders_${sheetIndex}`);
        const headerRowInput = document.getElementById(`headerRow_${sheetIndex}`);
        const dataRowInput = document.getElementById(`dataRow_${sheetIndex}`);
        
        if (hasHeadersCheckbox.checked) {
            // Enable header row input with defaults
            headerRowInput.disabled = false;
            headerRowInput.value = 1;
            dataRowInput.value = 2;
            dataRowInput.min = 2;
        } else {
            // Disable header row input and set data to start from row 1
            headerRowInput.disabled = true;
            headerRowInput.value = 0; // 0 indicates no headers
            dataRowInput.value = 1;
            dataRowInput.min = 1;
        }
    }
    
    detectRACIColumns(data) {
        // Expects data as array of arrays (rows)
        if (!data || data.length < 2) return [];
        
        const headers = data[0];
        const dataRows = data.slice(1);
        const raciPattern = /^[RACI,\s]+$/i; // Matches strings containing only R, A, C, I, commas, and spaces
        
        const columnTypes = headers.map((header, colIndex) => {
            let raciCount = 0;
            let totalNonEmpty = 0;
            
            // Sample up to first 20 rows to detect RACI pattern
            const sampleSize = Math.min(20, dataRows.length);
            
            for (let i = 0; i < sampleSize; i++) {
                const cellValue = dataRows[i][colIndex];
                if (cellValue !== undefined && cellValue !== null && cellValue !== '') {
                    totalNonEmpty++;
                    const stringValue = String(cellValue).trim();
                    
                    // Check if value matches RACI pattern
                    if (raciPattern.test(stringValue) || 
                        /^[RACI]$|^[RACI],[RACI]|^[RACI], [RACI]/.test(stringValue)) {
                        raciCount++;
                    }
                }
            }
            
            // If 50% or more of non-empty cells contain RACI values, mark as stakeholder column
            const isStakeholder = totalNonEmpty > 0 && (raciCount / totalNonEmpty) >= 0.5;
            
            return {
                name: header || `Column ${colIndex + 1}`,
                type: isStakeholder ? 'stakeholder' : 'information',
                index: colIndex
            };
        });
        
        return columnTypes;
    }
    
    toggleAllSheets(checked) {
        document.querySelectorAll('#sheetsList input[id^="sheet_"]').forEach(checkbox => {
            checkbox.checked = checked;
        });
    }
    
    processExcelImport() {
        if (!this.pendingExcelImport) return;
        
        const { workbook, fileName } = this.pendingExcelImport;
        
        // Get selected sheets with their configuration
        const selectedSheets = [];
        document.querySelectorAll('#sheetsList input[id^="sheet_"]:checked').forEach(checkbox => {
            const sheetIndex = parseInt(checkbox.value);
            const hasHeaders = document.getElementById(`hasHeaders_${sheetIndex}`).checked;
            const headerRow = parseInt(document.getElementById(`headerRow_${sheetIndex}`).value) || 1;
            const dataRow = parseInt(document.getElementById(`dataRow_${sheetIndex}`).value) || 2;
            
            // Validate row numbers
            if (hasHeaders && (headerRow < 1 || dataRow <= headerRow)) {
                alert(`Invalid configuration for sheet "${workbook.SheetNames[sheetIndex]}". Header row must be ≥ 1 and data row must be > header row.`);
                return;
            }
            if (!hasHeaders && dataRow < 1) {
                alert(`Invalid configuration for sheet "${workbook.SheetNames[sheetIndex]}". Data row must be ≥ 1.`);
                return;
            }
            
            selectedSheets.push({
                name: workbook.SheetNames[sheetIndex],
                index: sheetIndex,
                hasHeaders: hasHeaders,
                headerRowNumber: hasHeaders ? headerRow : 0,
                dataStartRowNumber: dataRow
            });
        });
        
        if (selectedSheets.length === 0) {
            alert('Please select at least one sheet to import.');
            return;
        }
        
        // Get import mode
        const importMode = document.querySelector('input[name="importMode"]:checked').value;
        
        try {
            const newPages = [];
            const usedNames = new Set(); // Track names used in this import batch
            const skippedSheets = []; // Track sheets that were skipped
            
            selectedSheets.forEach(sheetInfo => {
                const sheet = workbook.Sheets[sheetInfo.name];
                const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', blankrows: true });
                
                // Use per-sheet configuration
                const hasHeaders = sheetInfo.hasHeaders;
                const headerRowNumber = sheetInfo.headerRowNumber;
                const dataStartRowNumber = sheetInfo.dataStartRowNumber;
                
                // Check if we have enough rows based on configuration
                if (jsonData.length < dataStartRowNumber) {
                    console.warn(`Sheet "${sheetInfo.name}" has insufficient data (less than ${dataStartRowNumber} rows), skipping.`);
                    skippedSheets.push({ name: sheetInfo.name, reason: `Has only ${jsonData.length} rows, needs at least ${dataStartRowNumber}` });
                    return;
                }
                
                // Extract or generate headers
                let headers;
                if (hasHeaders) {
                    // Extract headers from the specified header row (convert to 0-based index)
                    const headerRowIndex = headerRowNumber - 1;
                    headers = jsonData[headerRowIndex];
                    
                    if (!headers || headers.length === 0) {
                        console.warn(`Sheet "${sheetInfo.name}" has no headers at row ${headerRowNumber}, skipping.`);
                        skippedSheets.push({ name: sheetInfo.name, reason: `No headers found at row ${headerRowNumber}` });
                        return;
                    }
                } else {
                    // Generate auto headers based on column count
                    const firstDataRow = jsonData[dataStartRowNumber - 1];
                    const colCount = firstDataRow ? firstDataRow.length : 0;
                    headers = Array.from({ length: colCount }, (_, i) => `Column ${String.fromCharCode(65 + i)}`);
                }
                
                // Extract data rows starting from the specified data row (convert to 0-based index)
                const dataStartIndex = dataStartRowNumber - 1;
                const allDataRows = jsonData.slice(dataStartIndex);
                
                // Prepare data for column detection (headers + data rows)
                const dataForDetection = [headers, ...allDataRows];
                
                // Detect column types
                const columnTypes = this.detectRACIColumns(dataForDetection);
                const dataRows = allDataRows.filter(row => row && row.some(cell => cell !== ''));
                
                // Create columns (skip first column as it will be activity name)
                const columns = [];
                for (let i = 1; i < columnTypes.length; i++) {
                    const colType = columnTypes[i];
                    columns.push({
                        id: this.generateId(),
                        name: colType.name,
                        type: colType.type,
                        description: ''
                    });
                }
                
                // Create rows
                const rows = [];
                dataRows.forEach(rowData => {
                    const activityName = rowData[0] ? String(rowData[0]).trim() : '';
                    if (!activityName) return; // Skip empty rows
                    
                    const cells = {};
                    columns.forEach((col, colIdx) => {
                        const cellValue = rowData[colIdx + 1]; // +1 because first column is activity name
                        
                        if (col.type === 'stakeholder') {
                            // Parse RACI values
                            const raciValues = [];
                            if (cellValue) {
                                const stringValue = String(cellValue).trim().toUpperCase();
                                // Extract R, A, C, I from the string
                                ['R', 'A', 'C', 'I'].forEach(letter => {
                                    if (stringValue.includes(letter)) {
                                        raciValues.push(letter);
                                    }
                                });
                            }
                            cells[col.id] = this.sortRACIValues(raciValues);
                        } else {
                            // Information column
                            cells[col.id] = cellValue ? [String(cellValue)] : [];
                        }
                    });
                    
                    rows.push({
                        id: this.generateId(),
                        name: activityName,
                        description: '',
                        cells: cells,
                        sectionId: null
                    });
                });
                
                // Create the page with unique name
                let pageName = sheetInfo.name;
                
                // Check for duplicates in existing project pages and in current import batch
                if (importMode === 'pages') {
                    pageName = this.getUniquePageNameWithBatch(pageName, usedNames);
                } else {
                    // Even for new projects, ensure unique names within the batch
                    pageName = this.getUniqueNameInBatch(pageName, usedNames);
                }
                
                usedNames.add(pageName);
                
                const page = {
                    id: this.generateId(),
                    name: pageName,
                    activityColumnName: headers[0] || 'Activity/Task',
                    columns: columns,
                    rows: rows,
                    sections: [],
                    searchQuery: ''
                };
                
                newPages.push(page);
            });
            
            // Apply import based on mode
            if (importMode === 'new') {
                // Create new project
                const projectName = fileName.replace(/\.[^/.]+$/, ''); // Remove file extension
                const newProject = {
                    id: this.generateId(),
                    projectName: projectName,
                    pages: newPages,
                    _currentPageIndex: 0
                };
                
                this.projects.push(newProject);
                this.currentProjectIndex = this.projects.length - 1;
                document.getElementById('projectName').value = this.project.projectName;
            } else {
                // Add pages to current project
                newPages.forEach(page => {
                    this.project.pages.push(page);
                });
            }
            
            this.render();
            this.saveToLocalStorage();
            this.updateProjectSwitcher();
            
            // Close modal and show success message
            this.modals.sheetSelection.hide();
            this.pendingExcelImport = null;
            
            let message = `Successfully imported ${newPages.length} sheet(s)!`;
            if (skippedSheets.length > 0) {
                message += `\n\nSkipped ${skippedSheets.length} sheet(s):`;
                skippedSheets.forEach(s => {
                    message += `\n- ${s.name}: ${s.reason}`;
                });
            }
            alert(message);
            
        } catch (error) {
            alert('Error processing Excel import: ' + error.message);
            console.error(error);
        }
    }

    async handleFileImport(event) {
        const file = event.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = async (e) => {
            try {
                let jsonString;
                const fileContent = e.target.result;
                
                // Check if it's an encrypted .raci file or plain .json
                if (file.name.endsWith('.raci')) {
                    // Decrypt the file
                    jsonString = await this.decryptData(fileContent);
                } else if (file.name.endsWith('.json')) {
                    // Plain JSON (backward compatibility)
                    jsonString = fileContent;
                } else {
                    // Try to detect format automatically
                    try {
                        // Try parsing as JSON first
                        JSON.parse(fileContent);
                        jsonString = fileContent;
                    } catch {
                        // If JSON parse fails, try decryption
                        jsonString = await this.decryptData(fileContent);
                    }
                }
                
                const imported = JSON.parse(jsonString);
                
                // Validate structure
                if (!imported.projectName || !imported.pages) {
                    alert('Invalid RACI project file');
                    return;
                }

                // Ensure backward compatibility - add activityColumnName if missing
                imported.pages.forEach(page => {
                    if (!page.activityColumnName) {
                        page.activityColumnName = 'Activity/Task';
                    }
                    if (!page.sections) {
                        page.sections = [];
                    }
                    // Add parentId to sections if missing (backward compatibility)
                    page.sections.forEach(section => {
                        if (section.parentId === undefined) {
                            section.parentId = null;
                        }
                    });
                    // Normalize RACI values in all cells to ensure correct order
                    page.rows.forEach(row => {
                        Object.keys(row.cells).forEach(colId => {
                            const col = page.columns.find(c => c.id === colId);
                            if (col && col.type === 'stakeholder' && Array.isArray(row.cells[colId])) {
                                row.cells[colId] = this.sortRACIValues([...row.cells[colId]]);
                            }
                        });
                    });
                });

                // Check if project with same name already exists
                const existingIndex = this.projects.findIndex(p => p.projectName === imported.projectName);
                
                if (existingIndex !== -1) {
                    // Project with same name exists - show modal dialog
                    this.pendingImport = {
                        data: imported,
                        existingIndex: existingIndex
                    };
                    
                    // Update modal with project name
                    document.getElementById('conflictProjectName').textContent = imported.projectName;
                    
                    // Show the modal
                    this.modals.importConflict.show();
                    return; // Exit here - handleImportConflict will complete the import
                } else {
                    // No conflict - add as new project
                    imported.id = this.generateId();
                    imported._currentPageIndex = 0;
                    this.projects.push(imported);
                    this.currentProjectIndex = this.projects.length - 1;
                }

                document.getElementById('projectName').value = this.project.projectName;
                this.render();
                this.saveToLocalStorage();
                this.updateProjectSwitcher();
                alert('Project imported successfully!');
            } catch (err) {
                alert('Error importing project: ' + err.message);
            }
        };
        reader.readAsText(file);
        event.target.value = ''; // Reset input
    }

    handleImportConflict(action) {
        // Hide the modal first
        this.modals.importConflict.hide();
        
        if (!this.pendingImport) {
            alert('No pending import found.');
            return;
        }
        
        const imported = this.pendingImport.data;
        const existingIndex = this.pendingImport.existingIndex;
        
        // Reset file input
        document.getElementById('fileInput').value = '';
        
        if (action === 'cancel') {
            // User cancelled - do nothing
            this.pendingImport = null;
            alert('Import cancelled.');
            return;
        }
        
        if (action === 'overwrite') {
            // Extra confirmation for overwrite
            const confirmOverwrite = confirm(
                `⚠️ WARNING: This will permanently delete the existing project "${imported.projectName}".\n\n` +
                `Are you absolutely sure you want to overwrite it?`
            );
            
            if (!confirmOverwrite) {
                this.pendingImport = null;
                alert('Import cancelled.');
                return;
            }
            
            // Overwrite existing project
            imported.id = this.projects[existingIndex].id; // Keep same ID
            imported._currentPageIndex = 0;
            this.projects[existingIndex] = imported;
            this.currentProjectIndex = existingIndex;
        } else if (action === 'new') {
            // Import as new project with modified name
            let newName = imported.projectName;
            let counter = 1;
            while (this.projects.some(p => p.projectName === newName)) {
                newName = `${imported.projectName} (${counter})`;
                counter++;
            }
            imported.projectName = newName;
            imported.id = this.generateId();
            imported._currentPageIndex = 0;
            this.projects.push(imported);
            this.currentProjectIndex = this.projects.length - 1;
        }
        
        // Complete the import
        document.getElementById('projectName').value = this.project.projectName;
        this.render();
        this.saveToLocalStorage();
        this.updateProjectSwitcher();
        this.pendingImport = null;
        
        alert('Project imported successfully!');
    }

    exportToExcel() {
        this.showMultiPageExportModal('excel');
    }
    
    performExcelExport(selectedIndices, useIndent = true) {
        const wb = XLSX.utils.book_new();

        selectedIndices.forEach(pageIndex => {
            const page = this.project.pages[pageIndex];
            // Prepare data
            const data = [];
            
            // Header row
            const activityColName = page.activityColumnName || 'Activity/Task';
            const header = [activityColName];
            page.columns.forEach(col => {
                header.push(col.name);
            });
            data.push(header);

            // Group rows by section
            const sections = page.sections || [];
            const rowsBySection = new Map();
            sections.forEach(s => rowsBySection.set(s.id, []));
            rowsBySection.set(null, []);
            page.rows.forEach(row => {
                const sectionId = row.sectionId || null;
                if (!rowsBySection.has(sectionId)) rowsBySection.set(sectionId, []);
                rowsBySection.get(sectionId).push(row);
            });

            // Track section rows with their levels for styling
            const sectionRowMetadata = [];

            // Render sections hierarchically
            const renderSectionAndRows = (sectionId, level = 0) => {
                const rows = rowsBySection.get(sectionId) || [];
                
                // Add section header if not null
                if (sectionId !== null && rows.length > 0) {
                    const section = sections.find(s => s.id === sectionId);
                    const indent = useIndent ? '  '.repeat(level) : '';
                    const sectionRow = [indent + section.name];
                    for (let i = 0; i < page.columns.length; i++) {
                        sectionRow.push('');
                    }
                    // Track this row index and its level
                    sectionRowMetadata.push({ rowIndex: data.length, level: level, sectionName: section.name });
                    data.push(sectionRow);
                }
                
                // Add rows with proper indentation
                rows.forEach(row => {
                    const indent = useIndent ? '  '.repeat(sectionId !== null ? level + 1 : 0) : '';
                    const rowData = [indent + row.name];
                    page.columns.forEach(col => {
                        const cellValue = row.cells[col.id] || [];
                        // Sort RACI values in correct order before joining
                        const sortedValue = col.type === 'stakeholder' ? 
                            this.sortRACIValues([...cellValue]) : cellValue;
                        rowData.push(sortedValue.join(', '));
                    });
                    data.push(rowData);
                });
                
                // Render child sections recursively
                if (sectionId !== null) {
                    const childSections = sections.filter(s => s.parentId === sectionId);
                    childSections.forEach(child => renderSectionAndRows(child.id, level + 1));
                }
            };
            
            // Start with "No Section" rows, then top-level sections
            renderSectionAndRows(null, 0);
            const topLevelSections = sections.filter(s => !s.parentId);
            topLevelSections.forEach(section => renderSectionAndRows(section.id, 0));

            // Create worksheet
            const ws = XLSX.utils.aoa_to_sheet(data);
            
            // Set column widths
            const colWidths = [{ wch: 60 }];
            page.columns.forEach(() => colWidths.push({ wch: 15 }));
            ws['!cols'] = colWidths;

            // Apply styling to cells
            const range = XLSX.utils.decode_range(ws['!ref']);
            
            for (let R = range.s.r; R <= range.e.r; ++R) {
                for (let C = range.s.c; C <= range.e.c; ++C) {
                    const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
                    if (!ws[cellAddress]) continue;
                    
                    // Initialize cell style
                    if (!ws[cellAddress].s) ws[cellAddress].s = {};
                    
                    // Header row styling (row 0)
                    if (R === 0) {
                        ws[cellAddress].s = {
                            fill: { fgColor: { rgb: "2980B9" } },
                            font: { bold: true, color: { rgb: "FFFFFF" }, sz: 12 },
                            alignment: { horizontal: "center", vertical: "center" },
                            border: {
                                top: { style: "thin", color: { rgb: "000000" } },
                                bottom: { style: "thin", color: { rgb: "000000" } },
                                left: { style: "thin", color: { rgb: "000000" } },
                                right: { style: "thin", color: { rgb: "000000" } }
                            }
                        };
                    } else {
                        // Check if this is a section header row and get its level
                        const sectionMeta = sectionRowMetadata.find(meta => meta.rowIndex === R);
                        
                        if (sectionMeta) {
                            // Define color scheme for different section levels
                            const levelColors = [
                                { bg: "D6EAF8", border: "2980B9" }, // Level 0: Light blue
                                { bg: "EBF5FB", border: "5DADE2" }, // Level 1: Very light blue
                                { bg: "F4F9FD", border: "85C1E9" }  // Level 2+: Almost white blue
                            ];
                            const colorIndex = Math.min(sectionMeta.level, levelColors.length - 1);
                            const colors = levelColors[colorIndex];
                            
                            ws[cellAddress].s = {
                                fill: { fgColor: { rgb: colors.bg } },
                                font: { bold: true, sz: 11 },
                                alignment: { horizontal: "left", vertical: "center" },
                                border: {
                                    left: { style: "medium", color: { rgb: colors.border } },
                                    top: { style: "thin", color: { rgb: "CCCCCC" } },
                                    bottom: { style: "thin", color: { rgb: "CCCCCC" } },
                                    right: { style: "thin", color: { rgb: "CCCCCC" } }
                                }
                            };
                        } else {
                            const col = page.columns[C - 1];
                            
                            // First column (Activity/Task)
                            if (C === 0) {
                                ws[cellAddress].s = {
                                    font: { bold: true, sz: 11 },
                                    alignment: { horizontal: "left", vertical: "center", wrapText: true },
                                    border: {
                                        top: { style: "thin", color: { rgb: "CCCCCC" } },
                                        bottom: { style: "thin", color: { rgb: "CCCCCC" } },
                                        left: { style: "thin", color: { rgb: "CCCCCC" } },
                                        right: { style: "thin", color: { rgb: "CCCCCC" } }
                                    }
                                };
                            } else if (col) {
                                // Stakeholder columns - light blue background
                                if (col.type === 'stakeholder') {
                                    ws[cellAddress].s = {
                                        fill: { fgColor: { rgb: "E3F2FD" } },
                                        alignment: { horizontal: "center", vertical: "center" },
                                        font: { sz: 11 },
                                        border: {
                                            top: { style: "thin", color: { rgb: "CCCCCC" } },
                                            bottom: { style: "thin", color: { rgb: "CCCCCC" } },
                                            left: { style: "thin", color: { rgb: "CCCCCC" } },
                                            right: { style: "thin", color: { rgb: "CCCCCC" } }
                                        }
                                    };
                                    
                                    // If cell has content, make it green
                                    if (ws[cellAddress].v) {
                                        ws[cellAddress].s.fill = { fgColor: { rgb: "D4EDDA" } };
                                        ws[cellAddress].s.font = { bold: true, color: { rgb: "155724" }, sz: 11 };
                                    }
                                } else {
                                    // Information columns - light teal background
                                    ws[cellAddress].s = {
                                        fill: { fgColor: { rgb: "E0F2F1" } },
                                        alignment: { horizontal: "left", vertical: "center" },
                                        font: { sz: 11 },
                                        border: {
                                            top: { style: "thin", color: { rgb: "CCCCCC" } },
                                            bottom: { style: "thin", color: { rgb: "CCCCCC" } },
                                            left: { style: "thin", color: { rgb: "CCCCCC" } },
                                            right: { style: "thin", color: { rgb: "CCCCCC" } }
                                        }
                                    };
                                }
                            }
                        }
                    }
                }
            }

            // Add worksheet to workbook with safe sheet name
            const sheetName = page.name.substring(0, 31).replace(/[:\\\/\?\*\[\]]/g, '_');
            XLSX.utils.book_append_sheet(wb, ws, sheetName);
        });

        // Save file
        XLSX.writeFile(wb, `${this.project.projectName || 'raci-matrix'}.xlsx`);
    }

    exportToPDF() {
        this.showMultiPageExportModal('pdf');
    }
    
    performPDFExport(selectedIndices, useIndent = true) {
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF('l', 'mm', 'a4'); // Landscape orientation

        let isFirstPage = true;

        selectedIndices.forEach(pageIndex => {
            const page = this.project.pages[pageIndex];
            if (!isFirstPage) {
                doc.addPage();
            }
            isFirstPage = false;

            // Title
            doc.setFontSize(16);
            doc.text(`${this.project.projectName} - ${page.name}`, 14, 15);

            // Prepare table data
            const activityColName = page.activityColumnName || 'Activity/Task';
            const headers = [[activityColName]];
            page.columns.forEach(col => {
                headers[0].push(col.name);
            });

            // Group rows by section
            const sections = page.sections || [];
            const rowsBySection = new Map();
            sections.forEach(s => rowsBySection.set(s.id, []));
            rowsBySection.set(null, []);
            page.rows.forEach(row => {
                const sectionId = row.sectionId || null;
                if (!rowsBySection.has(sectionId)) rowsBySection.set(sectionId, []);
                rowsBySection.get(sectionId).push(row);
            });

            const rows = [];
            const sectionRowMetadata = []; // Track section rows with their levels
            
            // Render sections hierarchically
            const renderSectionAndRows = (sectionId, level = 0) => {
                const sectionRows = rowsBySection.get(sectionId) || [];
                
                // Add section header if not null
                if (sectionId !== null && sectionRows.length > 0) {
                    const section = sections.find(s => s.id === sectionId);
                    const indent = useIndent ? '  '.repeat(level) : '';
                    const sectionRow = [indent + section.name];
                    for (let i = 0; i < page.columns.length; i++) {
                        sectionRow.push('');
                    }
                    sectionRowMetadata.push({ rowIndex: rows.length, level: level });
                    rows.push(sectionRow);
                }
                
                // Add data rows with proper indentation
                sectionRows.forEach(row => {
                    const indent = useIndent ? '  '.repeat(sectionId !== null ? level + 1 : 0) : '';
                    const rowData = [indent + row.name];
                    page.columns.forEach(col => {
                        const cellValue = row.cells[col.id] || [];
                        // Sort RACI values in correct order before joining
                        const sortedValue = col.type === 'stakeholder' ? 
                            this.sortRACIValues([...cellValue]) : cellValue;
                        rowData.push(sortedValue.join(', '));
                    });
                    rows.push(rowData);
                });
                
                // Render child sections recursively
                if (sectionId !== null) {
                    const childSections = sections.filter(s => s.parentId === sectionId);
                    childSections.forEach(child => renderSectionAndRows(child.id, level + 1));
                }
            };
            
            // Start with "No Section" rows, then top-level sections
            renderSectionAndRows(null, 0);
            const topLevelSections = sections.filter(s => !s.parentId);
            topLevelSections.forEach(section => renderSectionAndRows(section.id, 0));

            // Calculate dynamic column widths
            // A4 Landscape width: ~277mm, usable: ~250mm
            const totalWidth = 250;
            const stakeholderCols = page.columns.filter(c => c.type === 'stakeholder').length;
            const informationCols = page.columns.filter(c => c.type === 'information').length;
            
            const stakeholderWidth = stakeholderCols * 10; // 10% each
            const minInfoWidth = informationCols * 10; // minimum 10% each
            const infoRequiredPercent = (minInfoWidth / totalWidth) * 100;
            
            let activityWidth = 50; // Start with 50%
            let infoTotalWidth;
            
            if (infoRequiredPercent > 30) {
                // Borrow from activity column
                const borrowAmount = infoRequiredPercent - 30;
                activityWidth = 50 - borrowAmount;
                infoTotalWidth = 30;
            } else {
                infoTotalWidth = 100 - 50 - stakeholderWidth;
            }
            
            const infoWidthPerCol = informationCols > 0 ? (infoTotalWidth * totalWidth / 100) / informationCols : 0;
            
            // Build columnStyles object
            const columnStyles = {
                0: { 
                    cellWidth: (activityWidth * totalWidth / 100), 
                    fontStyle: 'bold',
                    halign: 'left'
                }
            };
            
            // Apply styles for each data column
            page.columns.forEach((col, idx) => {
                if (col.type === 'stakeholder') {
                    columnStyles[idx + 1] = {
                        cellWidth: (10 * totalWidth / 100),
                        halign: 'center',
                        fontStyle: 'bold'
                    };
                } else {
                    columnStyles[idx + 1] = {
                        cellWidth: infoWidthPerCol,
                        halign: 'left'
                    };
                }
            });

            // Generate table
            doc.autoTable({
                startY: 25,
                head: headers,
                body: rows,
                theme: 'grid',
                headStyles: { fillColor: [41, 128, 185], fontSize: 10, halign: 'center' },
                styles: { fontSize: 9, cellPadding: 3 },
                columnStyles: columnStyles,
                didParseCell: (data) => {
                    // Color code stakeholder vs information columns in header
                    if (data.row.section === 'head' && data.column.index > 0) {
                        const colIndex = data.column.index - 1;
                        if (page.columns[colIndex].type === 'information') {
                            data.cell.styles.fillColor = [52, 152, 219];
                        }
                    }
                    
                    // Style section header rows
                    if (data.row.section === 'body' && data.column.index === 0) {
                        const cellValue = data.cell.raw;
                        const isSection = sections.some(s => s.name === cellValue);
                        
                        if (isSection) {
                            data.cell.styles.fillColor = [233, 236, 239];
                            data.cell.styles.fontStyle = 'bold';
                            data.cell.styles.fontSize = 10;
                            data.cell.styles.textColor = [44, 62, 80];
                        }
                    }
                    
                    // Style section header rows based on level
                    if (data.row.section === 'body') {
                        const rowIndex = data.row.index - 1; // Adjust for header row
                        const sectionMeta = sectionRowMetadata.find(meta => meta.rowIndex === rowIndex);
                        
                        if (sectionMeta) {
                            // Define color scheme for different section levels (RGB format)
                            const levelColors = [
                                { bg: [214, 234, 248], text: [41, 128, 185] },   // Level 0: Light blue
                                { bg: [235, 245, 251], text: [93, 173, 226] },   // Level 1: Very light blue
                                { bg: [244, 249, 253], text: [133, 193, 233] }   // Level 2+: Almost white blue
                            ];
                            const colorIndex = Math.min(sectionMeta.level, levelColors.length - 1);
                            const colors = levelColors[colorIndex];
                            
                            data.cell.styles.fillColor = colors.bg;
                            data.cell.styles.textColor = colors.text;
                            data.cell.styles.fontStyle = 'bold';
                        }
                    }
                }
            });
        });

        doc.save(`${this.project.projectName || 'raci-matrix'}.pdf`);
    }

    exportToPPT() {
        this.showMultiPageExportModal('powerpoint');
    }
    
    performPowerPointExport(selectedIndices, useIndent = true) {
        const pptx = new PptxGenJS();

        selectedIndices.forEach(pageIndex => {
            const page = this.project.pages[pageIndex];
            // Prepare all table data first
            const allTableRows = [];
            
            // Header row
            const activityColName = page.activityColumnName || 'Activity/Task';
            const headerRow = [
                { text: activityColName, options: { bold: true, fill: '2980b9', color: 'FFFFFF', align: 'center' } }
            ];
            page.columns.forEach(col => {
                headerRow.push({
                    text: col.name,
                    options: {
                        bold: true,
                        fill: col.type === 'stakeholder' ? '2980b9' : '3498db',
                        color: 'FFFFFF',
                        align: 'center'
                    }
                });
            });

            // Group rows by section
            const sections = page.sections || [];
            const rowsBySection = new Map();
            sections.forEach(s => rowsBySection.set(s.id, []));
            rowsBySection.set(null, []);
            page.rows.forEach(row => {
                const sectionId = row.sectionId || null;
                if (!rowsBySection.has(sectionId)) rowsBySection.set(sectionId, []);
                rowsBySection.get(sectionId).push(row);
            });

            // Render sections hierarchically
            const renderSectionAndRows = (sectionId, level = 0) => {
                const sectionRowsData = rowsBySection.get(sectionId) || [];
                
                // Add section header if not null
                if (sectionId !== null && sectionRowsData.length > 0) {
                    const section = sections.find(s => s.id === sectionId);
                    const indent = useIndent ? '  '.repeat(level) : '';
                    const icon = level > 0 ? '↳' : '📁';
                    
                    // Define color scheme for different section levels
                    const levelStyles = [
                        { bg: 'D6EAF8', text: '2980B9' }, // Level 0: Light blue
                        { bg: 'EBF5FB', text: '5DADE2' }, // Level 1: Very light blue
                        { bg: 'F4F9FD', text: '85C1E9' }  // Level 2+: Almost white blue
                    ];
                    const styleIndex = Math.min(level, levelStyles.length - 1);
                    const style = levelStyles[styleIndex];
                    
                    const sectionRow = [
                        { text: `${indent}${icon} ${section.name}`, options: { bold: true, fill: style.bg, color: style.text, fontSize: 11, align: 'left' } }
                    ];
                    for (let i = 0; i < page.columns.length; i++) {
                        sectionRow.push({ text: '', options: { fill: style.bg } });
                    }
                    allTableRows.push(sectionRow);
                }
                
                // Add data rows with proper indentation
                sectionRowsData.forEach(row => {
                    const indent = useIndent ? '  '.repeat(sectionId !== null ? level + 1 : 0) : '';
                    const rowData = [{ text: indent + row.name, options: { bold: true, align: 'left' } }];
                    page.columns.forEach(col => {
                        const cellValue = row.cells[col.id] || [];
                        const cellOptions = {};
                        
                        // Configure based on column type
                        if (col.type === 'stakeholder') {
                            cellOptions.align = 'center';
                            cellOptions.bold = true;
                            // Add light background for stakeholder columns with data
                            if (cellValue.length > 0) {
                                cellOptions.fill = 'D4EDDA';
                                cellOptions.color = '155724';
                            }
                        } else {
                            // Information columns
                            cellOptions.align = 'left';
                        }
                        
                        // Sort RACI values in correct order before joining
                        const sortedValue = col.type === 'stakeholder' ? 
                            this.sortRACIValues([...cellValue]) : cellValue;
                        rowData.push({ text: sortedValue.join(', '), options: cellOptions });
                    });
                    allTableRows.push(rowData);
                });
                
                // Render child sections recursively
                if (sectionId !== null) {
                    const childSections = sections.filter(s => s.parentId === sectionId);
                    childSections.forEach(child => renderSectionAndRows(child.id, level + 1));
                }
            };
            
            // Start with "No Section" rows, then top-level sections
            renderSectionAndRows(null, 0);
            const topLevelSections = sections.filter(s => !s.parentId);
            topLevelSections.forEach(section => renderSectionAndRows(section.id, 0));
            
            // Calculate dynamic column widths for PowerPoint
            // Total slide width: 10 inches, usable: ~9 inches
            const totalWidth = 9.0;
            const stakeholderCols = page.columns.filter(c => c.type === 'stakeholder').length;
            const informationCols = page.columns.filter(c => c.type === 'information').length;
            
            const stakeholderWidthPercent = stakeholderCols * 10; // 10% each
            const minInfoWidthPercent = informationCols > 0 ? 30 : 0; // minimum 30% for info cols if any exist
            
            let activityWidthPercent = 50; // Start with 50%
            let infoTotalWidthPercent;
            
            if (informationCols > 0 && (stakeholderWidthPercent + 50 + minInfoWidthPercent > 100)) {
                // Need to borrow from activity column
                const totalNeeded = stakeholderWidthPercent + minInfoWidthPercent;
                activityWidthPercent = 100 - totalNeeded;
                infoTotalWidthPercent = minInfoWidthPercent;
            } else {
                infoTotalWidthPercent = 100 - 50 - stakeholderWidthPercent;
            }
            
            // Build column widths array
            const colWidths = [(activityWidthPercent / 100) * totalWidth];
            page.columns.forEach(col => {
                if (col.type === 'stakeholder') {
                    colWidths.push((10 / 100) * totalWidth);
                } else {
                    const infoWidthPerCol = informationCols > 0 ? (infoTotalWidthPercent / 100) * totalWidth / informationCols : 0;
                    colWidths.push(infoWidthPerCol);
                }
            });

            // Calculate rows per slide (accounting for title and margins)
            // Slide is 7.5" tall, title takes ~1", table starts at 1.0", so we have ~6.2" for content
            // With font size 10 and padding, each row is roughly 0.35", so ~17-18 rows max
            const MAX_ROWS_PER_SLIDE = 16; // Conservative to ensure good spacing
            
            // If all rows fit on one slide
            if (allTableRows.length <= MAX_ROWS_PER_SLIDE) {
                const slide = pptx.addSlide();
                
                // Title (y: 0.3 inches from top)
                slide.addText(`${this.project.projectName} - ${page.name}`, {
                    x: 0.5,
                    y: 0.3,
                    fontSize: 18,
                    bold: true,
                    color: '2980b9'
                });

                // Add complete table (y: 0.7 inches from top - reduced gap)
                slide.addTable([headerRow, ...allTableRows], {
                    x: 0.5,
                    y: 0.7,
                    colW: colWidths,
                    fontSize: 10,
                    border: { pt: 1, color: 'CFCFCF' },
                    valign: 'middle'
                });
            } else {
                // Split across multiple slides
                let currentSlideIndex = 1;
                let totalSlides = Math.ceil(allTableRows.length / MAX_ROWS_PER_SLIDE);
                
                for (let i = 0; i < allTableRows.length; i += MAX_ROWS_PER_SLIDE) {
                    const slide = pptx.addSlide();
                    
                    // Title with page indicator (y: 0.3 inches from top)
                    slide.addText(
                        `${this.project.projectName} - ${page.name} (${currentSlideIndex}/${totalSlides})`, 
                        {
                            x: 0.5,
                            y: 0.3,
                            fontSize: 18,
                            bold: true,
                            color: '2980b9'
                        }
                    );
                    
                    // Get rows for this slide
                    const slideRows = allTableRows.slice(i, i + MAX_ROWS_PER_SLIDE);
                    
                    // Add table with header row + data rows for this slide (y: 0.7 inches from top - reduced gap)
                    slide.addTable([headerRow, ...slideRows], {
                        x: 0.5,
                        y: 0.7,
                        colW: colWidths,
                        fontSize: 10,
                        border: { pt: 1, color: 'CFCFCF' },
                        valign: 'middle'
                    });
                    
                    currentSlideIndex++;
                }
            }
        });

        pptx.writeFile({ fileName: `${this.project.projectName || 'raci-matrix'}.pptx` });
    }

    exportToWord() {
        this.showMultiPageExportModal('word');
    }
    
    performWordExport(selectedIndices, useIndent = true) {
        // Generate HTML for Word document
        let htmlContent = `
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="UTF-8">
                <style>
                    body { font-family: Arial, sans-serif; }
                    h1 { color: #2980b9; font-size: 18pt; margin-bottom: 10pt; }
                    table { width: 100%; border-collapse: collapse; margin-bottom: 20pt; }
                    th { background-color: #2980b9; color: white; padding: 8pt; text-align: center; font-weight: bold; border: 1pt solid #cccccc; }
                    th.info-col { background-color: #3498db; }
                    td { padding: 6pt; border: 1pt solid #cccccc; }
                    td.activity-col { font-weight: bold; text-align: left; }
                    td.stakeholder-col { text-align: center; font-weight: bold; }
                    td.stakeholder-with-data { background-color: #d4edda; color: #155724; text-align: center; font-weight: bold; }
                    td.info-col { text-align: left; }
                    tr.section-header td { font-weight: bold; }
                    .indented { padding-left: 30pt; }
                </style>
            </head>
            <body>
        `;

        selectedIndices.forEach(pageIndex => {
            const page = this.project.pages[pageIndex];
            // Add page title
            htmlContent += `<h1>${this.project.projectName} - ${page.name}</h1>`;

            // Start table
            htmlContent += '<table>';

            // Header row
            const activityColName = page.activityColumnName || 'Activity/Task';
            htmlContent += '<thead><tr>';
            htmlContent += `<th>${activityColName}</th>`;
            page.columns.forEach(col => {
                const colClass = col.type === 'information' ? 'info-col' : '';
                htmlContent += `<th class="${colClass}">${col.name}</th>`;
            });
            htmlContent += '</tr></thead>';

            // Body rows
            htmlContent += '<tbody>';

            // Group rows by section
            const sections = page.sections || [];
            const rowsBySection = new Map();
            sections.forEach(s => rowsBySection.set(s.id, []));
            rowsBySection.set(null, []);
            page.rows.forEach(row => {
                const sectionId = row.sectionId || null;
                if (!rowsBySection.has(sectionId)) rowsBySection.set(sectionId, []);
                rowsBySection.get(sectionId).push(row);
            });

            // Render sections hierarchically
            const renderSectionAndRows = (sectionId, level = 0) => {
                const sectionRowsData = rowsBySection.get(sectionId) || [];

                // Add section header if not null
                if (sectionId !== null && sectionRowsData.length > 0) {
                    const section = sections.find(s => s.id === sectionId);
                    const indent = useIndent ? '&nbsp;&nbsp;'.repeat(level * 2) : '';
                    const icon = level > 0 ? '↳' : '📁';
                    
                    // Define color scheme for different section levels
                    const levelStyles = [
                        { bg: '#D6EAF8', color: '#2980B9' }, // Level 0: Light blue
                        { bg: '#EBF5FB', color: '#5DADE2' }, // Level 1: Very light blue
                        { bg: '#F4F9FD', color: '#85C1E9' }  // Level 2+: Almost white blue
                    ];
                    const styleIndex = Math.min(level, levelStyles.length - 1);
                    const style = levelStyles[styleIndex];
                    
                    htmlContent += `<tr class="section-header" style="background-color: ${style.bg};">`;
                    htmlContent += `<td colspan="${page.columns.length + 1}" style="font-weight: bold; color: ${style.color};">${indent}${icon} ${section.name}</td>`;
                    htmlContent += '</tr>';
                }

                // Add data rows with proper indentation
                sectionRowsData.forEach(row => {
                    const indentLevel = useIndent && sectionId !== null ? level + 1 : 0;
                    const indentStyle = indentLevel > 0 ? `padding-left: ${15 + indentLevel * 15}pt;` : '';
                    
                    htmlContent += '<tr>';
                    htmlContent += `<td class="activity-col" style="${indentStyle}">${row.name}</td>`;

                    page.columns.forEach(col => {
                        const cellValue = row.cells[col.id] || [];
                        // Sort RACI values in correct order before joining
                        const sortedValue = col.type === 'stakeholder' ? 
                            this.sortRACIValues([...cellValue]) : cellValue;
                        const cellText = sortedValue.join(', ');

                        let cellClass = '';
                        if (col.type === 'stakeholder') {
                            cellClass = cellValue.length > 0 ? 'stakeholder-with-data' : 'stakeholder-col';
                        } else {
                            cellClass = 'info-col';
                        }

                        htmlContent += `<td class="${cellClass}">${cellText}</td>`;
                    });

                    htmlContent += '</tr>';
                });
                
                // Render child sections recursively
                if (sectionId !== null) {
                    const childSections = sections.filter(s => s.parentId === sectionId);
                    childSections.forEach(child => renderSectionAndRows(child.id, level + 1));
                }
            };
            
            // Start with "No Section" rows, then top-level sections
            renderSectionAndRows(null, 0);
            const topLevelSections = sections.filter(s => !s.parentId);
            topLevelSections.forEach(section => renderSectionAndRows(section.id, 0));

            htmlContent += '</tbody></table>';
        });

        htmlContent += '</body></html>';

        // Convert HTML to DOCX
        const converted = htmlDocx.asBlob(htmlContent);
        saveAs(converted, `${this.project.projectName || 'raci-matrix'}.docx`);
    }

    // ========== UTILITIES ==========
    generateId() {
        return 'id_' + Math.random().toString(36).substr(2, 9) + Date.now().toString(36);
    }
    
    getUniquePageName(baseName) {
        // Check if a page with this name already exists in current project
        const existingNames = this.project.pages.map(p => p.name);
        
        if (!existingNames.includes(baseName)) {
            return baseName; // Name is unique, return as-is
        }
        
        // Name exists, find a unique suffix
        let counter = 2;
        let newName = `${baseName} (${counter})`;
        
        while (existingNames.includes(newName)) {
            counter++;
            newName = `${baseName} (${counter})`;
        }
        
        return newName;
    }
    
    getUniquePageNameWithBatch(baseName, usedNamesInBatch) {
        // Check against both existing pages and names used in current batch
        const existingNames = this.project.pages.map(p => p.name);
        const allNames = new Set([...existingNames, ...usedNamesInBatch]);
        
        if (!allNames.has(baseName)) {
            return baseName; // Name is unique
        }
        
        // Name exists, find a unique suffix
        let counter = 2;
        let newName = `${baseName} (${counter})`;
        
        while (allNames.has(newName)) {
            counter++;
            newName = `${baseName} (${counter})`;
        }
        
        return newName;
    }
    
    getUniqueNameInBatch(baseName, usedNamesInBatch) {
        // Check only against names used in current batch (for new projects)
        if (!usedNamesInBatch.has(baseName)) {
            return baseName; // Name is unique
        }
        
        // Name exists, find a unique suffix
        let counter = 2;
        let newName = `${baseName} (${counter})`;
        
        while (usedNamesInBatch.has(newName)) {
            counter++;
            newName = `${baseName} (${counter})`;
        }
        
        return newName;
    }
}

// Initialize app
const app = new RACIApp();

// Update project name on input change
document.getElementById('projectName').addEventListener('input', () => {
    app.updateProjectName();
});
