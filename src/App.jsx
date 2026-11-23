              <td>{r.unit_weight ?? 0}</td>
              <td>{weight(r.quantity, r.unit_weight)}</td>
              <td>{r.pre_assembly}</td>
              <td>{r.placement}</td>
              <td>{r.alignment}</td>
              <td>{r.bolting}</td>
              <td>{r.remark}</td>
              <td>
                <button
                  className="delete"
                  onClick={() => onDelete(r.id)}
                >
                  <Trash2 size={15} />
                </button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>

      {!rows.length && (
        <div className="empty">No records found.</div>
      )}
    </div>
  );
}

function DataTable({ rows, columns }) {
  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>
            {columns.map(c => (
              <th key={c}>{c.replaceAll('_', ' ')}</th>
            ))}
          </tr>
        </thead>

        <tbody>
          {rows.map((r, i) => (
            <tr key={r.id || i}>
              {columns.map(c => (
                <td key={c}>{r[c] ?? ''}</td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>

      {!rows.length && (
        <div className="empty">No records yet.</div>
      )}
    </div>
  );
}

function BillTable({ bills, type, paidFor, outstanding, onView }) {
  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Bill No.</th>
            <th>{type === 'vendor' ? 'Agency' : 'Project'}</th>
            <th>Bill Date</th>
            <th>Billing Period</th>
            <th>Bill Amount</th>
            <th>{type === 'vendor' ? 'Paid' : 'Received'}</th>
            <th>Outstanding</th>
            <th>View</th>
          </tr>
        </thead>

        <tbody>
          {bills.map(b => (
            <tr key={b.id}>
              <td><button className="link" onClick={() => onView({ bill: b, type })}>{b.bill_no}</button></td>
              <td>
                {type === 'vendor' ? b.agency_name : b.project_name}
              </td>
              <td>{b.bill_date}</td>
              <td>{b.billing_period}</td>
              <td>
                {money(
                  type === 'vendor'
                    ? (b.net_payable || b.total_amount)
                    : b.total_amount
                )}
              </td>
              <td>{money(paidFor(b.id))}</td>
              <td>{money(outstanding(b))}</td>
              <td>
                <button
                  className="secondary"
                  onClick={() => onView({ bill: b, type })}
                >
                  <Eye size={15} />
                  View
                </button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>

      {!bills.length && (
        <div className="empty">No bills saved yet.</div>
      )}
    </div>
  );
}

export default App;
