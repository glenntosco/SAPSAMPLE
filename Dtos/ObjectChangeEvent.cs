using System;
using System.Collections.Generic;

namespace Pro4Soft.SapB1Integration.Dtos
{
    public class ObjectChangeEvent
    {
        public Guid ObjectId { get; set; }
        public string Tenant { get; set; }
        public Guid? TenantId { get; set; }
        public string ReferenceNumber { get; set; }
        public Guid? UserId { get; set; }
        public Guid? UserSessionId { get; set; }
        public string ObjectType { get; set; }
        public DateTimeOffset Timestamp { get; set; } = DateTimeOffset.UtcNow;
        public EventType EventType { get; set; }
        public List<PropertyChange> Changes { get; set; } = new List<PropertyChange>();
    }

    public class PropertyChange
    {
        public string PropertyName { get; set; }
        public string ObjectType { get; set; }
        public object OldValue { get; set; }
        public object NewValue { get; set; }
    }

    public enum EventType
    {
        Added,
        Modified,
        Deleted
    }
}