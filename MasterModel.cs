var version = "$perspective"; // To Do: Replace this with the name of your perspective

// Remove tables, measures, columns and hierarchies that are not part of the perspective:
foreach(var t in Model.Tables.ToList()) {
    if(!t.InPerspective[version]) t.Delete();
    else {
        foreach(var m in t.Measures.ToList()) if(!m.InPerspective[version]) m.Delete();   
        foreach(var c in t.Columns.ToList()) if(!c.InPerspective[version]) c.Delete();
        foreach(var h in t.Hierarchies.ToList()) if(!h.InPerspective[version]) h.Delete();
    }
}

// Modify measures based on annotations:
foreach(Measure x in Model.AllMeasures) {
    var expr = x.GetAnnotation(version + "_Expression"); if(expr == null) continue;
    x.Expression = expr;
}

// ************* START UNHIDING ************** //
// Modify columns to unhide based on annotations:
foreach(Column x in Model.AllColumns) {
    var expr = x.GetAnnotation(version + "_Unhide"); if(expr == null) continue;
    x.IsHidden=false;
}

// Modify measures to unhide based on annotations
foreach(Measure x in Model.AllMeasures) {
    var expr = x.GetAnnotation(version + "_Unhide"); if(expr == null) continue;
    x.IsHidden=false;
}

// Modify tables to unhide based on annotations
foreach(Table x in Model.Tables.ToList()) {
    var expr = x.GetAnnotation(version + "_Unhide"); if(expr == null) continue;
    x.IsHidden=false;
}

// *************   END UNHIDING  ************** //

// ************* START HIDING ************** //
// Modify columns to unhide based on annotations:
foreach(Column x in Model.AllColumns) {
    var expr = x.GetAnnotation(version + "_Hide"); if(expr == null) continue;
    x.IsHidden=true;
}

// Modify measures to unhide based on annotations
foreach(Measure x in Model.AllMeasures) {
    var expr = x.GetAnnotation(version + "_Hide"); if(expr == null) continue;
    x.IsHidden=false;
}

// Modify tables to unhide based on annotations
foreach(Table x in Model.Tables.ToList()) {
    var expr = x.GetAnnotation(version + "_Hide"); if(expr == null) continue;
    x.IsHidden=true;
}

// *************   END HIDING  ************** //

// *************  START REMOVALS ************** //


// Remove Perspectives based on annotations
foreach(Perspective x in Model.Perspectives.ToList()) {
    var expr = x.GetAnnotation(version + "_Remove"); if(expr == null) continue;
    x.Delete();
}

// Remove Data Sources based on annotations
foreach(DataSource x in Model.DataSources.ToList()) {
    var expr = x.GetAnnotation(version + "_Remove"); if(expr == null) continue;
    x.Delete();
}

// Remove Roles based on annotations
foreach(ModelRole x in Model.Roles.ToList()) {
    var expr = x.GetAnnotation(version + "_Remove"); if(expr == null) continue;
    x.Delete();
}

// Remove Partitions based on annotations:
foreach(Table t in Model.Tables) {
    
    // Loop through all partitions in this table:
    foreach(Partition p in t.Partitions.ToList()) {
        var expr = p.GetAnnotation(version + "_Remove"); if(expr == null) continue;
        p.Delete();
    }
}

// Remove KPIs based on annotations
foreach(Measure x in Model.AllMeasures) {
    var expr = x.GetAnnotation(version + "_RemoveKPI"); if(expr == null) continue;
    x.KPI.Delete();
}
//   *************  END REMOVALS ************** //

// Set partition queries according to annotations:
foreach(Table t in Model.Tables) {
    var replaceValue = t.GetAnnotation(version + "_UpdatePartitionQuery"); if(replaceValue == null) continue;
    
    // Loop through all partitions in this table:
    foreach(Partition p in t.Partitions) {
        
        var finalQuery = p.Query.Replace("OldText", "NewText");

        // Replace all placeholder values:
        foreach(var placeholder in p.Annotations.Keys) {
            finalQuery = finalQuery.Replace("%" + placeholder + "%", p.GetAnnotation(placeholder));
        }

        p.Query = finalQuery;
    }
}

// Set partition data sources according to annotations:
foreach(Table t in Model.Tables) {
    var dataSourceName = t.GetAnnotation(version + "_UpdateDataSource"); if(dataSourceName == null) continue;
    
    // Loop through all partitions in this table:
    foreach(Partition p in t.Partitions) {
        p.DataSource = Model.DataSources[dataSourceName];
    }
}

// Update RLS based on annotations
foreach(Table t in Model.Tables) {
    foreach(ModelRole r in Model.Roles) {
        var rls = r.GetAnnotation(version + "_RLS_" + t.Name); if(rls == null) continue;
        t.RowLevelSecurity[r] = rls;
    }
}

// Update Role Members based on annotations
foreach(ModelRole r in Model.Roles) {
    var rm = r.GetAnnotation(version + "_UpdateRoleMembers"); if(rm == null) continue;
    r.RoleMembers = rm;
}

// Update Model Permissions from 'Administrator' to 'Read' based on annotations
    foreach(ModelRole r in Model.Roles) {
        var mp = r.GetAnnotation(version + "_ModelPermission"); if(mp == null) continue;
        if(mp == "Administrator") r.ModelPermission = ModelPermission.Administrator;
        else r.ModelPermission = ModelPermission.Read;
    }

