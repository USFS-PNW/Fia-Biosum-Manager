using System.Collections;
using System.Windows.Forms;

public class ListViewCompareChecked : IComparer
{
    private SortOrder m_SortOrder;

    public ListViewCompareChecked(SortOrder sort_order)
    {
        m_SortOrder = sort_order;
    }

    // Implements the IComparer interface
    public int Compare(object x, object y)
    {
        ListViewItem item_x = (ListViewItem)x;
        ListViewItem item_y = (ListViewItem)y;

        // The CompareTo method of the boolean type handles the comparison logic.
        // True (checked) is greater than False (unchecked).
        int compareResult = item_x.Checked.CompareTo(item_y.Checked);

        // Return the result based on the desired sort order.
        if (m_SortOrder == SortOrder.Ascending)
        {
            return compareResult;
        }
        else // Descending
        {
            return -compareResult;
        }
    }

    /// <summary>
    /// Gets or sets the order of sorting to apply (for example, 'Ascending' or 'Descending').
    /// </summary>
    public SortOrder Order
    {
        set
        {
            m_SortOrder = value;
        }
        get
        {
            return m_SortOrder;
        }
    }
}