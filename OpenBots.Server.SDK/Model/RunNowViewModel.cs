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
    /// RunNowViewModel
    /// </summary>
    [DataContract]
        public partial class RunNowViewModel :  IEquatable<RunNowViewModel>, IValidatableObject
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="RunNowViewModel" /> class.
        /// </summary>
        /// <param name="agentId">agentId.</param>
        /// <param name="agentGroupId">agentGroupId.</param>
        /// <param name="automationId">automationId (required).</param>
        /// <param name="jobParameters">jobParameters.</param>
        public RunNowViewModel(Guid? agentId = default(Guid?), Guid? agentGroupId = default(Guid?), Guid? automationId = default(Guid?), List<ParametersViewModel> jobParameters = default(List<ParametersViewModel>))
        {
            // to ensure "automationId" is required (not null)
            if (automationId == null)
            {
                throw new InvalidDataException("automationId is a required property for RunNowViewModel and cannot be null");
            }
            else
            {
                this.AutomationId = automationId;
            }
            this.AgentId = agentId;
            this.AgentGroupId = agentGroupId;
            this.JobParameters = jobParameters;
        }
        
        /// <summary>
        /// Gets or Sets AgentId
        /// </summary>
        [DataMember(Name="agentId", EmitDefaultValue=false)]
        public Guid? AgentId { get; set; }

        /// <summary>
        /// Gets or Sets AgentGroupId
        /// </summary>
        [DataMember(Name="agentGroupId", EmitDefaultValue=false)]
        public Guid? AgentGroupId { get; set; }

        /// <summary>
        /// Gets or Sets AutomationId
        /// </summary>
        [DataMember(Name="automationId", EmitDefaultValue=false)]
        public Guid? AutomationId { get; set; }

        /// <summary>
        /// Gets or Sets JobParameters
        /// </summary>
        [DataMember(Name="jobParameters", EmitDefaultValue=false)]
        public List<ParametersViewModel> JobParameters { get; set; }

        /// <summary>
        /// Returns the string presentation of the object
        /// </summary>
        /// <returns>String presentation of the object</returns>
        public override string ToString()
        {
            var sb = new StringBuilder();
            sb.Append("class RunNowViewModel {\n");
            sb.Append("  AgentId: ").Append(AgentId).Append("\n");
            sb.Append("  AgentGroupId: ").Append(AgentGroupId).Append("\n");
            sb.Append("  AutomationId: ").Append(AutomationId).Append("\n");
            sb.Append("  JobParameters: ").Append(JobParameters).Append("\n");
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
            return this.Equals(input as RunNowViewModel);
        }

        /// <summary>
        /// Returns true if RunNowViewModel instances are equal
        /// </summary>
        /// <param name="input">Instance of RunNowViewModel to be compared</param>
        /// <returns>Boolean</returns>
        public bool Equals(RunNowViewModel input)
        {
            if (input == null)
                return false;

            return 
                (
                    this.AgentId == input.AgentId ||
                    (this.AgentId != null &&
                    this.AgentId.Equals(input.AgentId))
                ) && 
                (
                    this.AgentGroupId == input.AgentGroupId ||
                    (this.AgentGroupId != null &&
                    this.AgentGroupId.Equals(input.AgentGroupId))
                ) && 
                (
                    this.AutomationId == input.AutomationId ||
                    (this.AutomationId != null &&
                    this.AutomationId.Equals(input.AutomationId))
                ) && 
                (
                    this.JobParameters == input.JobParameters ||
                    this.JobParameters != null &&
                    input.JobParameters != null &&
                    this.JobParameters.SequenceEqual(input.JobParameters)
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
                if (this.AgentId != null)
                    hashCode = hashCode * 59 + this.AgentId.GetHashCode();
                if (this.AgentGroupId != null)
                    hashCode = hashCode * 59 + this.AgentGroupId.GetHashCode();
                if (this.AutomationId != null)
                    hashCode = hashCode * 59 + this.AutomationId.GetHashCode();
                if (this.JobParameters != null)
                    hashCode = hashCode * 59 + this.JobParameters.GetHashCode();
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