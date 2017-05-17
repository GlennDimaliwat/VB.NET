Public Class Data
    Dim m_TICKET_SUMMARY As String
    Dim m_PROGRESS As String
    Dim m_CREATED As String
    Dim m_DUE_DATE As String
    Dim m_CATEGORY As String
    Dim m_SUB_CATEGORY As String
    Dim m_REQUESTOR As String
    Dim m_TICKET As String
    Dim m_DESCRIPTION As String
    Dim m_STATUS As String
    Dim m_REMARKS As String
    Dim m_CLASSIFICATION As String
    Dim m_ASSIGNEES As String
    Dim m_MODIFIED As String
    Dim m_RESOLVED As String
    Dim m_KB As String
    Dim m_DATE_LINKED_TO_KB As String
    Dim m_T_CODE As String
    Dim m_CHANGE_PHASE As String
    Dim m_REVIEW_DATE As String
    Dim m_AUTHORIZATION_DATE As String
    Dim m_BUILD_DATE As String
    Dim m_TESTING_DATE As String
    Dim m_IMPLEM_DATE As String
    Dim m_LAST_UPDATED As String

    Public Sub New(ByVal TICKET_SUMMARY As String,
                   ByVal PROGRESS As String,
                   ByVal CREATED As String,
                   ByVal DUE_DATE As String,
                   ByVal CATEGORY As String,
                   ByVal SUB_CATEGORY As String,
                   ByVal REQUESTOR As String,
                   ByVal TICKET As String,
                   ByVal DESCRIPTION As String,
                   ByVal STATUS As String,
                   ByVal REMARKS As String,
                   ByVal CLASSIFICATION As String,
                   ByVal ASSIGNEES As String,
                   ByVal MODIFIED As String,
                   ByVal RESOLVED As String,
                   ByVal KB As String,
                   ByVal DATE_LINKED_TO_KB As String,
                   ByVal T_CODE As String,
                   ByVal CHANGE_PHASE As String,
                   ByVal REVIEW_DATE As String,
                   ByVal AUTHORIZATION_DATE As String,
                   ByVal BUILD_DATE As String,
                   ByVal TESTING_DATE As String,
                   ByVal IMPLEM_DATE As String,
                   ByVal LAST_UPDATED As String)
        Me.m_TICKET_SUMMARY = TICKET_SUMMARY
        Me.m_PROGRESS = PROGRESS
        Me.m_CREATED = CREATED
        Me.m_DUE_DATE = DUE_DATE
        Me.m_CATEGORY = CATEGORY
        Me.m_SUB_CATEGORY = SUB_CATEGORY
        Me.m_REQUESTOR = REQUESTOR
        Me.m_TICKET = TICKET
        Me.m_DESCRIPTION = DESCRIPTION
        Me.m_STATUS = STATUS
        Me.m_REMARKS = REMARKS
        Me.m_CLASSIFICATION = CLASSIFICATION
        Me.m_ASSIGNEES = ASSIGNEES
        Me.m_MODIFIED = MODIFIED
        Me.m_KB = KB
        Me.m_DATE_LINKED_TO_KB = DATE_LINKED_TO_KB
        Me.m_T_CODE = T_CODE
        Me.m_CHANGE_PHASE = CHANGE_PHASE
        Me.m_REVIEW_DATE = REVIEW_DATE
        Me.m_AUTHORIZATION_DATE = AUTHORIZATION_DATE
        Me.m_BUILD_DATE = BUILD_DATE
        Me.m_TESTING_DATE = TESTING_DATE
        Me.m_IMPLEM_DATE = IMPLEM_DATE
        Me.m_RESOLVED = RESOLVED
        Me.m_LAST_UPDATED = LAST_UPDATED
    End Sub

    Property TICKET_SUMMARY() As String
        Get
            Return m_TICKET_SUMMARY
        End Get
        Set(ByVal value As String)
            m_TICKET_SUMMARY = value
        End Set
    End Property

    Property PROGRESS() As String
        Get
            Return m_PROGRESS
        End Get
        Set(ByVal value As String)
            m_PROGRESS = value
        End Set
    End Property

    Property CREATED() As String
        Get
            Return m_CREATED
        End Get
        Set(ByVal value As String)
            m_CREATED = value
        End Set
    End Property

    Property DUE_DATE() As String
        Get
            Return m_DUE_DATE
        End Get
        Set(ByVal value As String)
            m_DUE_DATE = value
        End Set
    End Property

    Property CATEGORY() As String
        Get
            Return m_CATEGORY
        End Get
        Set(ByVal value As String)
            m_CATEGORY = value
        End Set
    End Property

    Property SUB_CATEGORY() As String
        Get
            Return m_SUB_CATEGORY
        End Get
        Set(ByVal value As String)
            m_SUB_CATEGORY = value
        End Set
    End Property

    Property REQUESTOR() As String
        Get
            Return m_REQUESTOR
        End Get
        Set(ByVal value As String)
            m_REQUESTOR = value
        End Set
    End Property

    Property TICKET() As String
        Get
            Return m_TICKET
        End Get
        Set(ByVal value As String)
            m_TICKET = value
        End Set
    End Property

    Property DESCRIPTION() As String
        Get
            Return m_DESCRIPTION
        End Get
        Set(ByVal value As String)
            m_DESCRIPTION = value
        End Set
    End Property

    Property STATUS() As String
        Get
            Return m_STATUS
        End Get
        Set(ByVal value As String)
            m_STATUS = value
        End Set
    End Property

    Property REMARKS() As String
        Get
            Return m_REMARKS
        End Get
        Set(ByVal value As String)
            m_REMARKS = value
        End Set
    End Property

    Property CLASSIFICATION() As String
        Get
            Return m_CLASSIFICATION
        End Get
        Set(ByVal value As String)
            m_CLASSIFICATION = value
        End Set
    End Property

    Property ASSIGNEES() As String
        Get
            Return m_ASSIGNEES
        End Get
        Set(ByVal value As String)
            m_ASSIGNEES = value
        End Set
    End Property

    Property MODIFIED() As String
        Get
            Return m_MODIFIED
        End Get
        Set(ByVal value As String)
            m_MODIFIED = value
        End Set
    End Property

    Property RESOLVED() As String
        Get
            Return m_RESOLVED
        End Get
        Set(ByVal value As String)
            m_RESOLVED = value
        End Set
    End Property

    Property KB() As String
        Get
            Return m_KB
        End Get
        Set(ByVal value As String)
            m_KB = value
        End Set
    End Property

    Property DATE_LINKED_TO_KB() As String
        Get
            Return m_DATE_LINKED_TO_KB
        End Get
        Set(ByVal value As String)
            m_DATE_LINKED_TO_KB = value
        End Set
    End Property

    Property T_CODE() As String
        Get
            Return m_T_CODE
        End Get
        Set(ByVal value As String)
            m_T_CODE = value
        End Set
    End Property

    Property CHANGE_PHASE() As String
        Get
            Return m_CHANGE_PHASE
        End Get
        Set(ByVal value As String)
            m_CHANGE_PHASE = value
        End Set
    End Property

    Property REVIEW_DATE() As String
        Get
            Return m_REVIEW_DATE
        End Get
        Set(ByVal value As String)
            m_REVIEW_DATE = value
        End Set
    End Property

    Property AUTHORIZATION_DATE() As String
        Get
            Return m_AUTHORIZATION_DATE
        End Get
        Set(ByVal value As String)
            m_AUTHORIZATION_DATE = value
        End Set
    End Property

    Property BUILD_DATE() As String
        Get
            Return m_BUILD_DATE
        End Get
        Set(ByVal value As String)
            m_BUILD_DATE = value
        End Set
    End Property

    Property TESTING_DATE() As String
        Get
            Return m_TESTING_DATE
        End Get
        Set(ByVal value As String)
            m_TESTING_DATE = value
        End Set
    End Property

    Property IMPLEM_DATE() As String
        Get
            Return m_IMPLEM_DATE
        End Get
        Set(ByVal value As String)
            m_IMPLEM_DATE = value
        End Set
    End Property

    Property LAST_UPDATED As String
        Get
            Return m_LAST_UPDATED
        End Get
        Set(ByVal value As String)
            m_LAST_UPDATED = value
        End Set
    End Property
End Class
