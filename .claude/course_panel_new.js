      // ── Stats row ──
      stats.innerHTML =
        '<div style="flex:1;background:#ede9fe;border-radius:10px;padding:0.75rem 0.9rem;text-align:center;">' +
          '<div style="font-size:1.5rem;font-weight:800;color:#6366f1;">' + totalUpskilled + '</div>' +
          '<div style="font-size:0.68rem;color:#64748b;margin-top:2px;font-weight:500;">Teachers Upskilled</div>' +
        '</div>' +
        '<div style="flex:1;background:#dcfce7;border-radius:10px;padding:0.75rem 0.9rem;text-align:center;">' +
          '<div style="font-size:1.5rem;font-weight:800;color:#16a34a;">' + totalOnCourse + '</div>' +
          '<div style="font-size:0.68rem;color:#64748b;margin-top:2px;font-weight:500;">Active on Course</div>' +
        '</div>' +
        '<div style="flex:1;background:#fff7ed;border-radius:10px;padding:0.75rem 0.9rem;text-align:center;">' +
          '<div style="font-size:1.5rem;font-weight:800;color:#ea580c;">' + avgProf + '%</div>' +
          '<div style="font-size:0.68rem;color:#64748b;margin-top:2px;font-weight:500;">Avg Proficiency</div>' +
        '</div>';

      // ── Teacher list ──
      if (teachers.length === 0) {
        list.innerHTML = '<p style="text-align:center;color:var(--text-light);padding:2rem;">No teachers are upskilled on this course yet.</p>';
        return;
      }

      var rows = teachers.map(function(t, idx) {
        var prof      = parseFloat(t.proficiency) || 0;
        var profColor = prof >= 90 ? '#16a34a' : prof >= 71 ? '#ca8a04' : '#dc2626';
        var profBg    = prof >= 90 ? '#f0fdf4'  : prof >= 71 ? '#fffbeb' : '#fef2f2';
        var barWidth  = Math.min(Math.max(prof, 0), 100);
        var total     = t.totalLearners    || 0;
        var onCourse  = t.learnersOnCourse || 0;
        var bandwidth = total > 0 ? Math.round((onCourse / total) * 100) : 0;
        var bwColor   = bandwidth > 60 ? '#dc2626' : bandwidth > 30 ? '#ca8a04' : '#16a34a';
        var bwBg      = bandwidth > 60 ? '#fef2f2' : bandwidth > 30 ? '#fffbeb' : '#f0fdf4';
        var learners  = t.learners || [];
        var panelId   = 'cdl_' + idx;
        var initials  = escapeHtml((t.name || '').substring(0, 2).toUpperCase());

        var statusColors = {
          'Active Learner':   '#16a34a',
          'Friendly Learner': '#2563eb',
          'VIP':              '#7c3aed',
          'Break & Return':   '#d97706'
        };

        var learnerRows = learners.map(function(l) {
          var days    = l.daysOnCourse != null ? l.daysOnCourse + 'd' : '\u2014';
          var stCol   = statusColors[l.status] || '#64748b';
          var daysCol = (l.daysOnCourse || 0) > 180 ? '#b91c1c'
                      : (l.daysOnCourse || 0) > 90  ? '#d97706' : '#16a34a';
          return '<div style="display:flex;align-items:center;gap:8px;padding:6px 0;border-bottom:1px solid #f8fafc;">'
            + '<i class="fas fa-user-circle" style="color:#c7d2fe;font-size:0.88rem;flex-shrink:0;"></i>'
            + '<span style="flex:1;font-size:0.78rem;color:#374151;font-weight:500;">' + escapeHtml(l.name) + '</span>'
            + '<span style="font-size:0.7rem;font-weight:700;color:' + daysCol + ';background:' + bwBg + ';padding:2px 8px;border-radius:5px;min-width:32px;text-align:center;">' + days + '</span>'
            + '<span style="font-size:0.67rem;background:#f8fafc;color:' + stCol + ';padding:2px 9px;border-radius:5px;border:1px solid #e2e8f0;font-weight:600;white-space:nowrap;">' + escapeHtml(l.status || '\u2014') + '</span>'
            + '</div>';
        }).join('');

        // onclick toggle uses a data-target to avoid quoting hell
        var card =
          '<div style="background:#fff;border:1.5px solid #e2e8f0;border-radius:14px;margin-bottom:0.65rem;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,0.04);">'
          // Teacher header (clickable)
          + '<div class="cdl-header" data-panel="' + panelId + '" style="display:flex;align-items:center;gap:10px;padding:0.85rem 1rem;cursor:pointer;transition:background .12s;" onmouseover="this.style.background=\'#fafafa\'" onmouseout="this.style.background=\'\'">'
            + '<div style="width:38px;height:38px;border-radius:50%;background:linear-gradient(135deg,#6366f1,#8b5cf6);display:flex;align-items:center;justify-content:center;font-weight:800;font-size:0.88rem;color:#fff;flex-shrink:0;">' + initials + '</div>'
            + '<div style="flex:1;min-width:0;">'
              + '<div style="font-weight:700;font-size:0.88rem;color:#0f172a;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">' + escapeHtml(t.name) + '</div>'
              + '<div style="display:flex;align-items:center;gap:6px;margin-top:4px;">'
                + '<div style="background:#f1f5f9;border-radius:99px;height:4px;flex:1;overflow:hidden;max-width:80px;"><div style="background:' + profColor + ';height:4px;border-radius:99px;width:' + barWidth + '%;"></div></div>'
                + '<span style="font-size:0.7rem;font-weight:700;color:' + profColor + ';background:' + profBg + ';padding:1px 7px;border-radius:5px;">' + escapeHtml(t.proficiency || 'N/A') + '</span>'
              + '</div>'
            + '</div>'
            + '<div style="display:flex;flex-direction:column;align-items:flex-end;gap:4px;flex-shrink:0;">'
              + '<span style="font-size:0.78rem;font-weight:700;color:#6366f1;background:#ede9fe;padding:3px 10px;border-radius:8px;">' + onCourse + ' / ' + total + '</span>'
              + '<span style="font-size:0.65rem;color:#94a3b8;">BW: <strong style="color:' + bwColor + ';">' + bandwidth + '%</strong></span>'
            + '</div>'
            + '<i class="fas fa-chevron-right cdl-chevron" style="color:#cbd5e1;font-size:0.75rem;flex-shrink:0;transition:transform .2s;margin-left:4px;"></i>'
          + '</div>';

        if (learners.length > 0) {
          card += '<div id="' + panelId + '" style="display:none;padding:0 1rem 0.8rem;border-top:1px solid #f1f5f9;">'
            + '<div style="display:flex;gap:6px;padding:7px 0 6px;">'
              + '<span style="font-size:0.67rem;font-weight:600;color:#94a3b8;text-transform:uppercase;letter-spacing:.04em;flex:1;">Learner</span>'
              + '<span style="font-size:0.67rem;font-weight:600;color:#94a3b8;text-transform:uppercase;letter-spacing:.04em;min-width:38px;text-align:center;">Since</span>'
              + '<span style="font-size:0.67rem;font-weight:600;color:#94a3b8;text-transform:uppercase;letter-spacing:.04em;min-width:90px;text-align:center;">Stage</span>'
            + '</div>'
            + learnerRows
            + '</div>';
        } else {
          card += '<div style="padding:5px 1rem 0.65rem;font-size:0.74rem;color:#94a3b8;border-top:1px solid #f1f5f9;">No active learners on this course yet.</div>';
        }
        card += '</div>';
        return card;
      }).join('');

      list.innerHTML = '<div style="font-size:0.7rem;font-weight:600;color:#94a3b8;text-transform:uppercase;letter-spacing:.05em;margin-bottom:0.75rem;">'
        + teachers.length + ' Teacher' + (teachers.length !== 1 ? 's' : '') + ' \u2014 click to expand learners</div>'
        + rows;

      // Wire toggle clicks
      list.querySelectorAll('.cdl-header').forEach(function(hdr) {
        hdr.addEventListener('click', function() {
          var pid = this.getAttribute('data-panel');
          var panel = document.getElementById(pid);
          if (!panel) return;
          var open = panel.style.display !== 'none';
          panel.style.display = open ? 'none' : 'block';
          var ic = this.querySelector('.cdl-chevron');
          if (ic) ic.style.transform = open ? '' : 'rotate(90deg)';
        });
      });
    }
