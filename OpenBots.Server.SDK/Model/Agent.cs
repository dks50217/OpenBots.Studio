/* 
 * OpenBots Server API
 *
 * No description provided (generated by Swagger Codegen https://github.com/swagger-api/swagger-codegen)
 *
 * OpenAPI spec version: v1
 * 
 * Generated by: https://github.com/swagger-api/swagger-codegen.git
 */
using System;
using System.Linq;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Runtime.Serialization;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using System.ComponentModel.DataAnnotations;
using SwaggerDateConverter = OpenBots.Server.SDK.Client.SwaggerDateConverter;

namespace OpenBots.Server.SDK.Model
{
    /// <summary>
    /// Agent
    /// </summary>
    [DataContract]
        public partial class Agent :  IEquatable<Agent>, IValidatableObject
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="Agent" /> class.
        /// </summary>
        /// <param name="id">id.</param>
        /// <param name="isDeleted">isDeleted (default to false).</param>
        /// <param name="createdBy">createdBy.</param>
        /// <param name="createdOn">createdOn.</param>
        /// <param name="deletedBy">deletedBy.</param>
        /// <param name="deleteOn">deleteOn.</param>
        /// <param name="timestamp">timestamp.</param>
        /// <param name="updatedOn">updatedOn.</param>
        /// <param name="updatedBy">updatedBy.</param>
        /// <param name="name">name (required).</param>
        /// <param name="machineName">machineName (required).</param>
        /// <param name="macAddresses">macAddresses.</param>
        /// <param name="ipAddresses">ipAddresses.</param>
        /// <param name="isEnabled">isEnabled (required).</param>
        /// <param name="isConnected">isConnected (required).</param>
        /// <param name="credentialId">credentialId.</param>
        /// <param name="ipOption">ipOption.</param>
        /// <param name="isEnhancedSecurity">isEnhancedSecurity.</param>
        public Agent(Guid? id = default(Guid?), bool? isDeleted = false, string createdBy = default(string), DateTime? createdOn = default(DateTime?), string deletedBy = default(string), DateTime? deleteOn = default(DateTime?), byte[] timestamp = default(byte[]), DateTime? updatedOn = default(DateTime?), string updatedBy = default(string), string name = default(string), string machineName = default(string), string macAddresses = default(string), string ipAddresses = default(string), bool? isEnabled = default(bool?), bool? isConnected = default(bool?), Guid? credentialId = default(Guid?), string ipOption = default(string), bool? isEnhancedSecurity = default(bool?))
        {
            // to ensure "name" is required (not null)
            if (name == null)
            {
                throw new InvalidDataException("name is a required property for Agent and cannot be null");
            }
            else
            {
                this.Name = name;
            }
            // to ensure "machineName" is required (not null)
            if (machineName == null)
            {
                throw new InvalidDataException("machineName is a required property for Agent and cannot be null");
            }
            else
            {
                this.MachineName = machineName;
            }
            // to ensure "isEnabled" is required (not null)
            if (isEnabled == null)
            {
                throw new InvalidDataException("isEnabled is a required property for Agent and cannot be null");
            }
            else
            {
                this.IsEnabled = isEnabled;
            }
            // to ensure "isConnected" is required (not null)
            if (isConnected == null)
            {
                throw new InvalidDataException("isConnected is a required property for Agent and cannot be null");
            }
            else
            {
                this.IsConnected = isConnected;
            }
            this.Id = id;
            // use default value if no "isDeleted" provided
            if (isDeleted == null)
            {
                this.IsDeleted = false;
            }
            else
            {
                this.IsDeleted = isDeleted;
            }
            this.CreatedBy = createdBy;
            this.CreatedOn = createdOn;
            this.DeletedBy = deletedBy;
            this.DeleteOn = deleteOn;
            this.Timestamp = timestamp;
            this.UpdatedOn = updatedOn;
            this.UpdatedBy = updatedBy;
            this.MacAddresses = macAddresses;
            this.IpAddresses = ipAddresses;
            this.CredentialId = credentialId;
            this.IpOption = ipOption;
            this.IsEnhancedSecurity = isEnhancedSecurity;
        }
        
        /// <summary>
        /// Gets or Sets Id
        /// </summary>
        [DataMember(Name="id", EmitDefaultValue=false)]
        public Guid? Id { get; set; }

        /// <summary>
        /// Gets or Sets IsDeleted
        /// </summary>
        [DataMember(Name="isDeleted", EmitDefaultValue=false)]
        public bool? IsDeleted { get; set; }

        /// <summary>
        /// Gets or Sets CreatedBy
        /// </summary>
        [DataMember(Name="createdBy", EmitDefaultValue=false)]
        public string CreatedBy { get; set; }

        /// <summary>
        /// Gets or Sets CreatedOn
        /// </summary>
        [DataMember(Name="createdOn", EmitDefaultValue=false)]
        public DateTime? CreatedOn { get; set; }

        /// <summary>
        /// Gets or Sets DeletedBy
        /// </summary>
        [DataMember(Name="deletedBy", EmitDefaultValue=false)]
        public string DeletedBy { get; set; }

        /// <summary>
        /// Gets or Sets DeleteOn
        /// </summary>
        [DataMember(Name="deleteOn", EmitDefaultValue=false)]
        public DateTime? DeleteOn { get; set; }

        /// <summary>
        /// Gets or Sets Timestamp
        /// </summary>
        [DataMember(Name="timestamp", EmitDefaultValue=false)]
        public byte[] Timestamp { get; set; }

        /// <summary>
        /// Gets or Sets UpdatedOn
        /// </summary>
        [DataMember(Name="updatedOn", EmitDefaultValue=false)]
        public DateTime? UpdatedOn { get; set; }

        /// <summary>
        /// Gets or Sets UpdatedBy
        /// </summary>
        [DataMember(Name="updatedBy", EmitDefaultValue=false)]
        public string UpdatedBy { get; set; }

        /// <summary>
        /// Gets or Sets Name
        /// </summary>
        [DataMember(Name="name", EmitDefaultValue=false)]
        public string Name { get; set; }

        /// <summary>
        /// Gets or Sets MachineName
        /// </summary>
        [DataMember(Name="machineName", EmitDefaultValue=false)]
        public string MachineName { get; set; }

        /// <summary>
        /// Gets or Sets MacAddresses
        /// </summary>
        [DataMember(Name="macAddresses", EmitDefaultValue=false)]
        public string MacAddresses { get; set; }

        /// <summary>
        /// Gets or Sets IpAddresses
        /// </summary>
        [DataMember(Name="ipAddresses", EmitDefaultValue=false)]
        public string IpAddresses { get; set; }

        /// <summary>
        /// Gets or Sets IsEnabled
        /// </summary>
        [DataMember(Name="isEnabled", EmitDefaultValue=false)]
        public bool? IsEnabled { get; set; }

        /// <summary>
        /// Gets or Sets IsConnected
        /// </summary>
        [DataMember(Name="isConnected", EmitDefaultValue=false)]
        public bool? IsConnected { get; set; }

        /// <summary>
        /// Gets or Sets CredentialId
        /// </summary>
        [DataMember(Name="credentialId", EmitDefaultValue=false)]
        public Guid? CredentialId { get; set; }

        /// <summary>
        /// Gets or Sets IpOption
        /// </summary>
        [DataMember(Name="ipOption", EmitDefaultValue=false)]
        public string IpOption { get; set; }

        /// <summary>
        /// Gets or Sets IsEnhancedSecurity
        /// </summary>
        [DataMember(Name="isEnhancedSecurity", EmitDefaultValue=false)]
        public bool? IsEnhancedSecurity { get; set; }

        /// <summary>
        /// Returns the string presentation of the object
        /// </summary>
        /// <returns>String presentation of the object</returns>
        public override string ToString()
        {
            var sb = new StringBuilder();
            sb.Append("class Agent {\n");
            sb.Append("  Id: ").Append(Id).Append("\n");
            sb.Append("  IsDeleted: ").Append(IsDeleted).Append("\n");
            sb.Append("  CreatedBy: ").Append(CreatedBy).Append("\n");
            sb.Append("  CreatedOn: ").Append(CreatedOn).Append("\n");
            sb.Append("  DeletedBy: ").Append(DeletedBy).Append("\n");
            sb.Append("  DeleteOn: ").Append(DeleteOn).Append("\n");
            sb.Append("  Timestamp: ").Append(Timestamp).Append("\n");
            sb.Append("  UpdatedOn: ").Append(UpdatedOn).Append("\n");
            sb.Append("  UpdatedBy: ").Append(UpdatedBy).Append("\n");
            sb.Append("  Name: ").Append(Name).Append("\n");
            sb.Append("  MachineName: ").Append(MachineName).Append("\n");
            sb.Append("  MacAddresses: ").Append(MacAddresses).Append("\n");
            sb.Append("  IpAddresses: ").Append(IpAddresses).Append("\n");
            sb.Append("  IsEnabled: ").Append(IsEnabled).Append("\n");
            sb.Append("  IsConnected: ").Append(IsConnected).Append("\n");
            sb.Append("  CredentialId: ").Append(CredentialId).Append("\n");
            sb.Append("  IpOption: ").Append(IpOption).Append("\n");
            sb.Append("  IsEnhancedSecurity: ").Append(IsEnhancedSecurity).Append("\n");
            sb.Append("}\n");
            return sb.ToString();
        }
  
        /// <summary>
        /// Returns the JSON string presentation of the object
        /// </summary>
        /// <returns>JSON string presentation of the object</returns>
        public virtual string ToJson()
        {
            return JsonConvert.SerializeObject(this, Formatting.Indented);
        }

        /// <summary>
        /// Returns true if objects are equal
        /// </summary>
        /// <param name="input">Object to be compared</param>
        /// <returns>Boolean</returns>
        public override bool Equals(object input)
        {
            return this.Equals(input as Agent);
        }

        /// <summary>
        /// Returns true if Agent instances are equal
        /// </summary>
        /// <param name="input">Instance of Agent to be compared</param>
        /// <returns>Boolean</returns>
        public bool Equals(Agent input)
        {
            if (input == null)
                return false;

            return 
                (
                    this.Id == input.Id ||
                    (this.Id != null &&
                    this.Id.Equals(input.Id))
                ) && 
                (
                    this.IsDeleted == input.IsDeleted ||
                    (this.IsDeleted != null &&
                    this.IsDeleted.Equals(input.IsDeleted))
                ) && 
                (
                    this.CreatedBy == input.CreatedBy ||
                    (this.CreatedBy != null &&
                    this.CreatedBy.Equals(input.CreatedBy))
                ) && 
                (
                    this.CreatedOn == input.CreatedOn ||
                    (this.CreatedOn != null &&
                    this.CreatedOn.Equals(input.CreatedOn))
                ) && 
                (
                    this.DeletedBy == input.DeletedBy ||
                    (this.DeletedBy != null &&
                    this.DeletedBy.Equals(input.DeletedBy))
                ) && 
                (
                    this.DeleteOn == input.DeleteOn ||
                    (this.DeleteOn != null &&
                    this.DeleteOn.Equals(input.DeleteOn))
                ) && 
                (
                    this.Timestamp == input.Timestamp ||
                    (this.Timestamp != null &&
                    this.Timestamp.Equals(input.Timestamp))
                ) && 
                (
                    this.UpdatedOn == input.UpdatedOn ||
                    (this.UpdatedOn != null &&
                    this.UpdatedOn.Equals(input.UpdatedOn))
                ) && 
                (
                    this.UpdatedBy == input.UpdatedBy ||
                    (this.UpdatedBy != null &&
                    this.UpdatedBy.Equals(input.UpdatedBy))
                ) && 
                (
                    this.Name == input.Name ||
                    (this.Name != null &&
                    this.Name.Equals(input.Name))
                ) && 
                (
                    this.MachineName == input.MachineName ||
                    (this.MachineName != null &&
                    this.MachineName.Equals(input.MachineName))
                ) && 
                (
                    this.MacAddresses == input.MacAddresses ||
                    (this.MacAddresses != null &&
                    this.MacAddresses.Equals(input.MacAddresses))
                ) && 
                (
                    this.IpAddresses == input.IpAddresses ||
                    (this.IpAddresses != null &&
                    this.IpAddresses.Equals(input.IpAddresses))
                ) && 
                (
                    this.IsEnabled == input.IsEnabled ||
                    (this.IsEnabled != null &&
                    this.IsEnabled.Equals(input.IsEnabled))
                ) && 
                (
                    this.IsConnected == input.IsConnected ||
                    (this.IsConnected != null &&
                    this.IsConnected.Equals(input.IsConnected))
                ) && 
                (
                    this.CredentialId == input.CredentialId ||
                    (this.CredentialId != null &&
                    this.CredentialId.Equals(input.CredentialId))
                ) && 
                (
                    this.IpOption == input.IpOption ||
                    (this.IpOption != null &&
                    this.IpOption.Equals(input.IpOption))
                ) && 
                (
                    this.IsEnhancedSecurity == input.IsEnhancedSecurity ||
                    (this.IsEnhancedSecurity != null &&
                    this.IsEnhancedSecurity.Equals(input.IsEnhancedSecurity))
                );
        }

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Hash code</returns>
        public override int GetHashCode()
        {
            unchecked // Overflow is fine, just wrap
            {
                int hashCode = 41;
                if (this.Id != null)
                    hashCode = hashCode * 59 + this.Id.GetHashCode();
                if (this.IsDeleted != null)
                    hashCode = hashCode * 59 + this.IsDeleted.GetHashCode();
                if (this.CreatedBy != null)
                    hashCode = hashCode * 59 + this.CreatedBy.GetHashCode();
                if (this.CreatedOn != null)
                    hashCode = hashCode * 59 + this.CreatedOn.GetHashCode();
                if (this.DeletedBy != null)
                    hashCode = hashCode * 59 + this.DeletedBy.GetHashCode();
                if (this.DeleteOn != null)
                    hashCode = hashCode * 59 + this.DeleteOn.GetHashCode();
                if (this.Timestamp != null)
                    hashCode = hashCode * 59 + this.Timestamp.GetHashCode();
                if (this.UpdatedOn != null)
                    hashCode = hashCode * 59 + this.UpdatedOn.GetHashCode();
                if (this.UpdatedBy != null)
                    hashCode = hashCode * 59 + this.UpdatedBy.GetHashCode();
                if (this.Name != null)
                    hashCode = hashCode * 59 + this.Name.GetHashCode();
                if (this.MachineName != null)
                    hashCode = hashCode * 59 + this.MachineName.GetHashCode();
                if (this.MacAddresses != null)
                    hashCode = hashCode * 59 + this.MacAddresses.GetHashCode();
                if (this.IpAddresses != null)
                    hashCode = hashCode * 59 + this.IpAddresses.GetHashCode();
                if (this.IsEnabled != null)
                    hashCode = hashCode * 59 + this.IsEnabled.GetHashCode();
                if (this.IsConnected != null)
                    hashCode = hashCode * 59 + this.IsConnected.GetHashCode();
                if (this.CredentialId != null)
                    hashCode = hashCode * 59 + this.CredentialId.GetHashCode();
                if (this.IpOption != null)
                    hashCode = hashCode * 59 + this.IpOption.GetHashCode();
                if (this.IsEnhancedSecurity != null)
                    hashCode = hashCode * 59 + this.IsEnhancedSecurity.GetHashCode();
                return hashCode;
            }
        }

        /// <summary>
        /// To validate all properties of the instance
        /// </summary>
        /// <param name="validationContext">Validation context</param>
        /// <returns>Validation Result</returns>
        IEnumerable<System.ComponentModel.DataAnnotations.ValidationResult> IValidatableObject.Validate(ValidationContext validationContext)
        {
            yield break;
        }
    }
}