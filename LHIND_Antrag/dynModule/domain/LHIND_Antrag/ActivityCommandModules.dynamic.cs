// This file is part of Acta Nova (www.acta-nova.eu)
// Copyright (c) rubicon IT GmbH, www.rubicon.eu
// Version 1.4 - Philipp Rössler - 25.10.2019

using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using ActaNova.Domain;
using ActaNova.Domain.Classes.Configuration;
using ActaNova.Domain.Classifications;
using ActaNova.Domain.Extensions;
using ActaNova.Domain.Specialdata;
using ActaNova.Domain.Specialdata.Values;
using ActaNova.Domain.Workflow;
using Remotion.Data.DomainObjects;
using Remotion.Globalization;
using Remotion.Logging;
using Remotion.ObjectBinding;
using Remotion.SecurityManager.Domain.OrganizationalStructure;
using Rubicon.Dms;
using Rubicon.Domain;
using Rubicon.Domain.TenantBoundOrganizationalStructure;
using Rubicon.Gever.Bund.Domain.Utilities.Extensions;
using Rubicon.Gever.Bund.EGovInterface.Domain.Export;
using Rubicon.Gever.Bund.EGovInterface.Domain.Import;
using Rubicon.Gever.Bund.EGovInterface.Domain.Import.Mapping;
using Rubicon.Gever.Bund.EGovInterface.Domain.Reader;
using Rubicon.Multilingual.Extensions;
using Rubicon.ObjectBinding.Utilities;
using Rubicon.Utilities;
using Rubicon.Utilities.Autofac;
using Rubicon.Utilities.Globalization;
using Rubicon.Utilities.Security;
using Rubicon.Workflow.Domain;
using Document = ActaNova.Domain.Document;

namespace LHIND.Antrag
{
    [LocalizationEnum]
    public enum LocalizedUserMessages
    {
        [MultiLingualName("Das Geschäftsobjekt der Aktivität \"{0}\" muss ein Dossier sein.", ""),
         MultiLingualName("L’objet métier de l’activité \"{0}\" doit être un dossier.", "Fr"),
         MultiLingualName("L'oggetto business dell'attività \"{0}\" deve essere un dossier.", "It")]
        NoFile,

        [MultiLingualName("Das Geschäftsobjekt der Aktivität \"{0}\" muss ein Geschäftsvorfall sein.", ""),
         MultiLingualName("L’objet métier de l’activité \"{0}\" doit être une opération d’affaire.", "Fr"),
         MultiLingualName("L'oggetto business dell'attività \"{0}\" deve essere un'operazione business.", "It")]
        NoFileCase,

        [MultiLingualName("Das Geschäftsobjekt \"{0}\" der Aktivität kann nicht bearbeitet werden.", ""),
         MultiLingualName("L’objet métier \"{0}\" de l’activité ne peut être édité.", "Fr"),
         MultiLingualName("L'oggetto business \"{0}\" dell'attività non può essere modificato.", "It")]
        NotEditable,

        [MultiLingualName("Es wurde keine Geschäftsobjektvorlage mit der Referenz-ID \"{0}\" gefunden.", ""),
         MultiLingualName("Aucun modèle d’objets métier avec le numéro de référence \"{0}\" n’a été trouvé.", "Fr"),
         MultiLingualName("Non è stato trovato alcun modello di oggetto business con l'ID di riferimento.", "It")]
        NoFileCaseTemplate,

        [MultiLingualName("Es wurde noch keine Federführung Amt definiert.", ""),
         MultiLingualName("Aucune unité d’organisation responsable n’a été définie.", "Fr"),
         MultiLingualName("Non è stato ancora definito un ufficio responsabile.", "It")]
        NoFfDefined,

        [MultiLingualName("Es wurde keine Gruppe für den Katalogwert \"{0}\" gefunden.", ""),
         MultiLingualName("Es wurde keine Gruppe fär den Katalogwert \"{0}\" gefunden.", "Fr"),
         MultiLingualName("Non è stato trovato nessun gruppo per il valore del catalogo \"{0}\".", "It")]
        NoGroupDefined,

        [MultiLingualName("Es wurde keine Geschäftsart \"{0}\" gefunden.", ""),
         MultiLingualName("Aucun groupe pour la valeur de catalogue \"{0}\" n’a été trouvé.", "Fr"),
         MultiLingualName("Non è stato trovato nessun gruppo per il valore del catalogo \"{0}\".", "It")]
        NoFileCaseType,

        [MultiLingualName("Es wurden keine Empfänger \"{0}\" gefunden.", ""),
         MultiLingualName("Aucun type d’affaire \"{0}\" n’a été trouvé.", "Fr"),
         MultiLingualName("Non è stato trovato nessun destinatario \"{0}\".", "It")]
        NoRecipients,

        [MultiLingualName("Das eCH-Paket wurde noch nicht importiert.", ""),
         MultiLingualName("Le paquet n'a pas encore été importé.", "Fr"),
         MultiLingualName("Il pacco non è stato ancora consegnato.", "It")]
        FileNotImported,

        [MultiLingualName("Es konnte kein Eingang mit Referenz zum Quellobjekt im Dossier (\"{0}\") gefunden werden.", ""),
         MultiLingualName("...", "Fr"),
         MultiLingualName("...", "It")]
        IncomingNotFound
    }

    public class LhindAntragCreateInternalFilecases : ActivityCommandModule
    {
        private static readonly ILog s_logger = LogManager.GetLogger(typeof(LhindAntragCreateInternalFilecases));

        public LhindAntragCreateInternalFilecases() : base("LHIND_Antrag_GvfErzeugen:ActivityCommandClassificationType")
        {
        }

        public override bool Execute(CommandActivity commandActivity)
        {
            //Set templates for FileCase creation.
            const string fileCaseFfTemplateReferenceId = "9B9A5D57-3DC5-41A6-81DB-6D4ADC1335E6";
            const string fileCaseMbTemplateReferenceId = "7AAA0700-CA11-4759-B250-DF3BC05B9754";

            var startFileCases = new List<FileCase>();

            try
            {
                var hostObject = (File) commandActivity.WorkItem;

                if (hostObject == null)
                    throw new ActivityCommandException("").WithUserMessage(LocalizedUserMessages.NoFile
                        .ToLocalizedName()
                        .FormatWith(commandActivity));

                if (!hostObject.CanEdit(true))
                    throw new ActivityCommandException("").WithUserMessage(LocalizedUserMessages.NotEditable
                        .ToLocalizedName()
                        .FormatWith(hostObject));

                //Create FF FileCase
                var fileCaseTemplateFf =
                    new ReferenceHandle<FileCaseTemplate>(fileCaseFfTemplateReferenceId).GetObject();
                if (fileCaseTemplateFf == null)
                    throw new ActivityCommandException("").WithUserMessage(LocalizedUserMessages.NoFileCaseTemplate
                        .ToLocalizedName()
                        .FormatWith(fileCaseFfTemplateReferenceId));
                var newTitle = String.Empty;
                var ffGroup = hostObject.GetProperty("#LHIND_Antrag_Ff") as TenantGroup;

                if (ffGroup != null)
                {
                    var ffFileCase = FileCase.NewObject(hostObject, null, null, fileCaseTemplateFf);

                    ffFileCase.LeadingGroup = ffGroup;

                    newTitle = ffFileCase.GetMultilingualValue(fc => fc.Title) + " - " +
                               ffGroup.GetMultilingualValue(g => g.ShortName);
                    ffFileCase.SetMultilingualValue(fc => fc.Title, newTitle);

                    startFileCases.Add(ffFileCase);
                }

                //Create MB FileCases
                var fileCaseTemplateMb =
                    new ReferenceHandle<FileCaseTemplate>(fileCaseMbTemplateReferenceId).GetObject();
                if (fileCaseTemplateMb == null)
                    throw new ActivityCommandException("").WithUserMessage(LocalizedUserMessages.NoFileCaseTemplate
                        .ToLocalizedName()
                        .FormatWith(fileCaseMbTemplateReferenceId));

                var mbFileCaseRecipients =
                    hostObject.GetProperty("#LHIND_Antrag_Mb") as SpecialdataListPropertyValueCollection;

                foreach (var fileCaseRecipient in mbFileCaseRecipients)
                {
                    var mbFileCase = FileCase.NewObject(hostObject, null, null, fileCaseTemplateMb);
                    mbFileCase.LeadingGroup = fileCaseRecipient.Unwrap() as TenantGroup;

                    newTitle = mbFileCase.GetMultilingualValue(fc => fc.Title) + " - " +
                               mbFileCase.LeadingGroup.GetMultilingualValue(g => g.ShortName);
                    mbFileCase.SetMultilingualValue(fc => fc.Title, newTitle);

                    startFileCases.Add(mbFileCase);
                }

                ClientTransaction.Current.Commit();

                foreach (var startFileCase in startFileCases) startFileCase.StartObject();

                ClientTransaction.Current.Commit();
            }
            catch (Exception ex)
            {
                s_logger.Error(ex.Message, ex);
                ClientTransaction.Current.Rollback();
                throw;
            }

            return true;
        }

        public override string Validate(CommandActivity commandActivity)
        {
            if (commandActivity.EffectiveWorkItemType.IsAssignableFrom(typeof(File)))
                return null;
            else
                return "Invalid WorkItemType";
        }
    }

    public class LhindAntragCreateIncoming : ActivityCommandModule
    {
        private static readonly ILog s_logger = LogManager.GetLogger(typeof(LhindAntragCreateIncoming));

        public LhindAntragCreateIncoming() : base("LHIND_Antrag_EingangErstellen:ActivityCommandClassificationType")
        {
        }

        public override bool Execute(CommandActivity commandActivity)
        {
            const string targetIncomingTypeReferenceId = "7CFEE47F-72C7-4DDF-B2C1-71CA03E1A808";

            var restoreCtxId = ApplicationContext.CurrentID;

            try
            {
                var hostObject = (FileCase) commandActivity.WorkItem;

                if (hostObject == null)
                    throw new ActivityCommandException("").WithUserMessage(LocalizedUserMessages.NoFileCase
                        .ToLocalizedName()
                        .FormatWith(commandActivity));

                if (!hostObject.CanEdit(true))
                    throw new ActivityCommandException("").WithUserMessage(LocalizedUserMessages.NotEditable
                        .ToLocalizedName()
                        .FormatWith(hostObject));

                var sourceFileCaseUrl = UrlProvider.Current.GetOpenWorkListItemUrl(hostObject);

                //Create eCH-0147 container
                var messageExport = Containers.Global.Resolve<IMessageExport>();
                var eCHExport = messageExport.Export(hostObject);

                hostObject.AddFileCaseContent(eCHExport);

                //switch tenant
                using (ClientTransaction.CreateRootTransaction().EnterDiscardingScope())
                using (TenantSection.SwitchToTenant(UserHelper.Current.Tenant))
                {
                    //Create new incoming, set specialdata properties
                    var incoming = Incoming.NewObject();
                    ApplicationContext.CurrentID = incoming.ApplicationContextID;

                    incoming.Subject = hostObject.DisplayName + " - eCH-Dossier";
                    incoming.IncomingType =
                        new ReferenceHandle<IncomingTypeClassificationType>(targetIncomingTypeReferenceId).GetObject();
                    incoming.ExternalNumber = hostObject.FormattedNumber;
                    incoming.Remark = hostObject.WorkInstruction;
                    using (new SpecialdataIgnoreReadOnlySection())
                        incoming.SetProperty("#LHIND_Antrag_sourceFileCaseUrl", sourceFileCaseUrl);

                    var targeteCHDocument = Document.NewObject(incoming);
                    ((IDocument) targeteCHDocument).Name = hostObject.GetMultilingualValue(fc => fc.Title) + " (" +
                                                           hostObject.FormattedNumber + ") - eCH Import";
                    targeteCHDocument.PhysicallyPresent = false;
                    targeteCHDocument.Type = (DocumentClassificationType) ClassificationType.GetObject(Rubicon.Gever
                        .Bund.EGovInterface.Domain.WellKnownObjects.DocumentClassification.EchImport.GetObjectID());

                    using (TenantSection.DisableQueryRestrictions())
                    using (var handle = eCHExport.ActiveContent.GetContent())
                        targeteCHDocument.ActiveContent.SetContent(handle, "zip", "application/zip");

                    var targetFile = ImportHelper.TenantKnowsObject(targeteCHDocument, true);
                    if (targetFile != null)
                    {
                        incoming.LeadingGroup = targetFile.LeadingGroup;
                        incoming.Insert(targetFile);
                    }
                    else
                    {
                        incoming.LeadingGroup = UserHelper.Current.GetActaNovaUserExtension().StandardGroup != null
                            ? UserHelper.Current.GetActaNovaUserExtension().StandardGroup.AsTenantGroup()
                            : UserHelper.Current.OwningGroup.AsTenantGroup();
                    }

                    ClientTransaction.Current.Commit();
                }
            }
            catch (Exception ex)
            {
                s_logger.Error(ex.Message, ex);
                throw;
            }
            finally
            {
                ApplicationContext.CurrentID = restoreCtxId;
            }

            return true;
        }

        public override string Validate(CommandActivity commandActivity)
        {
            if (commandActivity.EffectiveWorkItemType.IsAssignableFrom(typeof(FileCase)))
                return null;
            else
                return "Invalid WorkItemType";
        }
    }

    public class LhindAntragRegisterIncomingToeCHImport : ActivityCommandModule
    {
        private static readonly ILog s_logger = LogManager.GetLogger(typeof(LhindAntragRegisterIncomingToeCHImport));

        public LhindAntragRegisterIncomingToeCHImport() : base(
            "LHIND_Antrag_eCHEingangRegistrieren:ActivityCommandClassificationType")
        {
        }

        public override bool Execute(CommandActivity commandActivity)
        {
            try
            {
                var hostObject = (Incoming) commandActivity.WorkItem;

                if (hostObject.BaseFile != null) return true;

                //Get final Dossier
                var echDocument = hostObject.FileContentHierarchyFlat.OfType<Document>()
                    .First(d => d.Type == ClassificationType.GetObject(Rubicon.Gever.Bund.EGovInterface.Domain
                                    .WellKnownObjects.DocumentClassification.EchImport.GetObjectID()));

                var targetFile = ImportHelper.TenantKnowsObject(echDocument);

                if (targetFile == null)
                    throw new ActivityCommandException("").WithUserMessage(LocalizedUserMessages.FileNotImported
                        .ToLocalizedName()
                        .FormatWith(commandActivity));

                hostObject.Insert(targetFile);
            }
            catch (Exception ex)
            {
                s_logger.Error(ex.Message, ex);
                throw;
            }

            return true;
        }

        public override string Validate(CommandActivity commandActivity)
        {
            if (commandActivity.EffectiveWorkItemType.IsAssignableFrom(typeof(Incoming)))
                return null;
            else
                return "Invalid WorkItemType";
        }
    }

    public class LhindAntragReturnToSender : ActivityCommandModule
    {
        private static readonly ILog s_logger = LogManager.GetLogger(typeof(LhindAntragReturnToSender));

        public LhindAntragReturnToSender() : base("LHIND_Antrag_Ruecksendung:ActivityCommandClassificationType")
        {
        }

        public override bool Execute(CommandActivity commandActivity)
        {
            var transferIncomingTypeReferenceId = "7CFEE47F-72C7-4DDF-B2C1-71CA03E1A808";

            var restoreCtxId = ApplicationContext.CurrentID;

            try
            {
                var sourceFileCase = (FileCase) commandActivity.WorkItem;

                if (sourceFileCase == null)
                    throw new ActivityCommandException("").WithUserMessage(LocalizedUserMessages.NoFileCase
                        .ToLocalizedName()
                        .FormatWith(commandActivity));

                if (!sourceFileCase.CanEdit(true))
                    throw new ActivityCommandException("").WithUserMessage(LocalizedUserMessages.NotEditable
                        .ToLocalizedName()
                        .FormatWith(sourceFileCase));

                var sourceIncoming = sourceFileCase.BaseFile.BaseIncomings
                    .Where(i => i.IncomingType.ToHasReferenceID().ReferenceID.ToUpper() ==
                                transferIncomingTypeReferenceId.ToUpper())
                    .OrderByDescending(t => t.CreatedAt)
                    .FirstOrDefault();
                if (sourceIncoming == null)
                    throw new ActivityCommandException("").WithUserMessage(LocalizedUserMessages.IncomingNotFound
                        .ToLocalizedName()
                        .FormatWith(sourceFileCase.BaseFile));

                var targetFileCaseUri = new Uri(sourceIncoming.GetProperty("#LHIND_Antrag_sourceFileCaseUrl") as string);

                var targetTenantId = HttpUtility.ParseQueryString(targetFileCaseUri.Query).Get("TenantID");

                var sourceFileCaseUrl = UrlProvider.Current.GetOpenWorkListItemUrl(sourceFileCase);

                //Create eCH-0147 container
                var messageExport = Containers.Global.Resolve<IMessageExport>();
                var eCHExport = messageExport.Export(sourceFileCase);

                sourceFileCase.AddFileCaseContent(eCHExport);

                //Switch tenant
                using (ClientTransaction.CreateRootTransaction().EnterDiscardingScope())
                using (TenantSection.SwitchToTenant(Tenant.FindByUnqiueIdentifier(targetTenantId)))
                {
                    //Create new incoming in target tenant
                    var incoming = Incoming.NewObject();
                    ApplicationContext.CurrentID = incoming.ApplicationContextID;

                    incoming.Subject = sourceFileCase.DisplayName + " - eCH Response";
                    incoming.IncomingType =
                        new ReferenceHandle<IncomingTypeClassificationType>(transferIncomingTypeReferenceId)
                            .GetObject();
                    incoming.ExternalNumber = sourceFileCase.FormattedNumber;
                    incoming.Remark = sourceFileCase.WorkInstruction;
                    using (new SpecialdataIgnoreReadOnlySection())
                        incoming.SetProperty("#LHIND_Antrag_sourceFileCaseUrl", sourceFileCaseUrl);

                    var targeteCHDocument = Document.NewObject(incoming);
                    ((IDocument) targeteCHDocument).Name = sourceFileCase.GetMultilingualValue(fc => fc.Title) + " (" +
                                                           sourceFileCase.FormattedNumber + ") - eCH Import";
                    targeteCHDocument.PhysicallyPresent = false;
                    targeteCHDocument.Type = (DocumentClassificationType) ClassificationType.GetObject(Rubicon.Gever
                        .Bund.EGovInterface.Domain.WellKnownObjects.DocumentClassification.EchImport.GetObjectID());

                    using (TenantSection.DisableQueryRestrictions())
                    using (var handle = eCHExport.ActiveContent.GetContent())
                        targeteCHDocument.ActiveContent.SetContent(handle, "zip", "application/zip");

                    var targetFile = ImportHelper.TenantKnowsObject(targeteCHDocument, true);
                    if (targetFile != null)
                    {
                        incoming.LeadingGroup = targetFile.LeadingGroup;
                        incoming.Insert(targetFile);
                    }
                    else
                    {
                        incoming.LeadingGroup = UserHelper.Current.GetActaNovaUserExtension().StandardGroup != null
                            ? UserHelper.Current.GetActaNovaUserExtension().StandardGroup.AsTenantGroup()
                            : UserHelper.Current.OwningGroup.AsTenantGroup();
                    }

                    ClientTransaction.Current.Commit();
                }
            }
            catch (Exception ex)
            {
                s_logger.Error(ex.Message, ex);
                throw;
            }
            finally
            {
                ApplicationContext.CurrentID = restoreCtxId;
            }

            return true;
        }

        public override string Validate(CommandActivity commandActivity)
        {
            if (commandActivity.EffectiveWorkItemType.IsAssignableFrom(typeof(FileCase)))
                return null;
            else
                return "Invalid WorkItemType";
        }
    }

    static class ImportHelper
    {
        public static File TenantKnowsObject(Document eCHDocument, bool autoImport = false)
        {
            using (var handle = eCHDocument.ActiveContent.GetContent())
            using (var echZipReader = new StreamEchZipReader(handle.CreateStream()))
            {
                var message = echZipReader.ExtractMessage();
                if (message != null & message.content != null && message.content.dossiers != null &&
                    message.content.dossiers.dossier != null && message.content.dossiers.dossier.Count == 1)
                {
                    var file = MessageImportHelper.TryGetExistingDomainObject<File>(message.content.dossiers.dossier
                        .First()
                        .applicationCustom);

                    if (file != null && autoImport)
                    {
                        var messageImport = MessageImport.NewObject(Containers.Global.Resolve<IMessageImportMapper>(),
                            echZipReader, MessageImportModeType.Values.CreateAndUpdate());
                        messageImport.ExecuteImport();
                    }

                    return file;
                }
            }

            return null;
        }
    }
}