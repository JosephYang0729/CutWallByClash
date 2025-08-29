using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using CutWallByClash.Models;
using System;
using System.Collections.Generic;

namespace CutWallByClash.Core
{
    public class WallOpeningExternalEvent : IExternalEventHandler
    {
        private List<ClashInfo> _clashes;
        private OpeningParameters _parameters;
        private IProgress<int> _progress;
        private Action<string, Dictionary<ClashInfo, ElementId>> _onCompleted;
        private Action<string> _onError;

        public void SetData(List<ClashInfo> clashes, OpeningParameters parameters, 
            IProgress<int> progress, Action<string, Dictionary<ClashInfo, ElementId>> onCompleted, Action<string> onError)
        {
            _clashes = clashes;
            _parameters = parameters;
            _progress = progress;
            _onCompleted = onCompleted;
            _onError = onError;
        }

        public void Execute(UIApplication app)
        {
            try
            {
                var document = app.ActiveUIDocument.Document;
                var openingCreator = new OpeningCreator(document);

                _progress?.Report(10);

                // 群組碰撞點
                var groups = openingCreator.GroupClashes(_clashes, _parameters);
                
                _progress?.Report(30);

                // 在 ExternalEvent 中創建 Transaction
                using (var transaction = new Transaction(document, "Create Wall Openings"))
                {
                    transaction.Start();

                    try
                    {
                        // 創建開口並獲取開口ID
                        var openingIds = openingCreator.CreateOpenings(groups, new Progress<int>(p => 
                        {
                            _progress?.Report(30 + (int)(p * 0.7)); // 30% + 70% of creation progress
                        }));

                        transaction.Commit();
                        
                        _progress?.Report(100);
                        _onCompleted?.Invoke($"開口建立完成！處理了 {groups.Count} 個開口群組。", openingIds);
                    }
                    catch (Exception ex)
                    {
                        transaction.RollBack();
                        throw;
                    }
                }
            }
            catch (Exception ex)
            {
                _onError?.Invoke($"建立開口時發生錯誤: {ex.Message}");
            }
        }

        public string GetName()
        {
            return "Wall Opening External Event";
        }
    }

    public class WallOpeningEventHandler
    {
        private readonly WallOpeningExternalEvent _externalEvent;
        private readonly ExternalEvent _revitEvent;

        public WallOpeningEventHandler()
        {
            _externalEvent = new WallOpeningExternalEvent();
            _revitEvent = ExternalEvent.Create(_externalEvent);
        }

        public void ExecuteOpeningCreation(List<ClashInfo> clashes, OpeningParameters parameters,
            IProgress<int> progress, Action<string, Dictionary<ClashInfo, ElementId>> onCompleted, Action<string> onError)
        {
            _externalEvent.SetData(clashes, parameters, progress, onCompleted, onError);
            _revitEvent.Raise();
        }
    }
}